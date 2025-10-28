// QueueManager.gs - キュー管理システム

/**
 * ファイルキュー管理、優先度ベースの処理順制御、並列処理制限の管理を行うクラス
 * バッチ処理と単一ファイル処理の両方に対応
 */
class QueueManager {
  constructor() {
    this.cache = CacheService.getUserCache();
    this.maxConcurrentTasks = CONFIG.BATCH_PROCESSING.PARALLEL_PROCESSING_LIMIT;
    this.maxRetryAttempts = CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS;
    this.queueKey = 'translation_queue';
    this.activeTasksKey = 'active_tasks';
  }

  /**
   * キューを初期化する 
   * @return {Object} 初期化結果
   */
  initializeQueue() {
    try {
      const queue = {
        tasks: [],
        activeTasks: {},
        completedTasks: [],
        failedTasks: [],
        createdAt: new Date().getTime(),
        lastUpdated: new Date().getTime()
      };

      this.cache.put(this.queueKey, JSON.stringify(queue), 21600); // 6時間有効
      log('INFO', 'キューを初期化しました');

      return { status: 'initialized', queue: queue };

    } catch (error) {
      log('ERROR', 'キュー初期化エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * タスクをキューに追加する
   * @param {Object} taskData - タスクデータ
   * @param {number} priority - 優先度（1-10、10が最高優先度）
   * @return {Object} 追加結果
   */
  enqueueTask(taskData, priority = 5) {
    try {
      const queue = this.getQueue() || this.initializeQueue().queue;
      
      const task = {
        taskId: taskData.taskId || this.generateTaskId(),
        type: taskData.type || 'single_file', // 'single_file' または 'batch'
        priority: Math.max(1, Math.min(10, priority)), // 1-10の範囲に制限
        status: 'queued',
        data: taskData,
        createdAt: new Date().getTime(),
        updatedAt: new Date().getTime(),
        retryCount: 0,
        errorMessage: '',
        estimatedDuration: taskData.estimatedDuration || 0
      };

      // 優先度順で挿入位置を決定
      let insertIndex = queue.tasks.length;
      for (let i = 0; i < queue.tasks.length; i++) {
        if (queue.tasks[i].priority < task.priority) {
          insertIndex = i;
          break;
        }
      }

      queue.tasks.splice(insertIndex, 0, task);
      queue.lastUpdated = new Date().getTime();

      this.saveQueue(queue);

      log('INFO', `タスクをキューに追加: ${task.taskId}, 優先度=${priority}, 位置=${insertIndex + 1}/${queue.tasks.length}`);

      return {
        status: 'enqueued',
        taskId: task.taskId,
        queuePosition: insertIndex + 1,
        totalTasks: queue.tasks.length
      };

    } catch (error) {
      log('ERROR', 'タスクキュー追加エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * 次に処理すべきタスクを取得する
   * @return {Object|null} 次のタスク
   */
  dequeueTask() {
    try {
      const queue = this.getQueue();
      if (!queue || queue.tasks.length === 0) {
        return null;
      }

      // アクティブタスク数をチェック
      const activeTaskCount = Object.keys(queue.activeTasks).length;
      if (activeTaskCount >= this.maxConcurrentTasks) {
        log('DEBUG', `並列処理制限に達しました: ${activeTaskCount}/${this.maxConcurrentTasks}`);
        return null;
      }

      // 最高優先度のタスクを取得
      const task = queue.tasks.shift();
      if (!task) {
        return null;
      }

      // アクティブタスクとして登録
      task.status = 'processing';
      task.startedAt = new Date().getTime();
      task.updatedAt = new Date().getTime();
      
      queue.activeTasks[task.taskId] = task;
      queue.lastUpdated = new Date().getTime();

      this.saveQueue(queue);

      log('INFO', `タスクを処理開始: ${task.taskId}, タイプ=${task.type}, 優先度=${task.priority}`);

      return task;

    } catch (error) {
      log('ERROR', 'タスクデキューエラー', error);
      return null;
    }
  }

  /**
   * タスクの処理を完了としてマーク
   * @param {string} taskId - タスクID
   * @param {Object} result - 処理結果
   * @return {Object} 完了処理結果
   */
  completeTask(taskId, result) {
    try {
      const queue = this.getQueue();
      if (!queue || !queue.activeTasks[taskId]) {
        return { status: 'error', message: 'アクティブタスクが見つかりません' };
      }

      const task = queue.activeTasks[taskId];
      task.status = 'completed';
      task.completedAt = new Date().getTime();
      task.updatedAt = new Date().getTime();
      task.result = result;
      task.duration = (task.completedAt - task.startedAt) / 1000;

      // アクティブタスクから完了タスクに移動
      delete queue.activeTasks[taskId];
      queue.completedTasks.push(task);
      queue.lastUpdated = new Date().getTime();

      this.saveQueue(queue);

      log('INFO', `タスク完了: ${taskId}, 処理時間=${Math.round(task.duration)}秒`);

      return {
        status: 'completed',
        taskId: taskId,
        duration: task.duration,
        result: result
      };

    } catch (error) {
      log('ERROR', 'タスク完了処理エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * タスクを失敗としてマークし、必要に応じて再キューイング
   * @param {string} taskId - タスクID
   * @param {string} errorMessage - エラーメッセージ
   * @param {boolean} shouldRetry - 再試行するかどうか
   * @return {Object} 失敗処理結果
   */
  failTask(taskId, errorMessage, shouldRetry = true) {
    try {
      const queue = this.getQueue();
      if (!queue || !queue.activeTasks[taskId]) {
        return { status: 'error', message: 'アクティブタスクが見つかりません' };
      }

      const task = queue.activeTasks[taskId];
      task.errorMessage = errorMessage;
      task.retryCount++;
      task.updatedAt = new Date().getTime();

      // 再試行判定
      if (shouldRetry && task.retryCount <= this.maxRetryAttempts) {
        // 再キューイング（優先度を少し下げる）
        task.status = 'queued';
        task.priority = Math.max(1, task.priority - 1);
        
        // アクティブタスクから通常のキューに戻す
        delete queue.activeTasks[taskId];
        
        // 優先度順で再挿入
        let insertIndex = queue.tasks.length;
        for (let i = 0; i < queue.tasks.length; i++) {
          if (queue.tasks[i].priority < task.priority) {
            insertIndex = i;
            break;
          }
        }
        
        queue.tasks.splice(insertIndex, 0, task);
        queue.lastUpdated = new Date().getTime();

        this.saveQueue(queue);

        log('WARN', `タスク再キューイング: ${taskId}, 再試行回数=${task.retryCount}/${this.maxRetryAttempts}, 新優先度=${task.priority}`);

        return {
          status: 'requeued',
          taskId: taskId,
          retryCount: task.retryCount,
          queuePosition: insertIndex + 1
        };

      } else {
        // 最大再試行回数に達した、または再試行不要
        task.status = 'failed';
        task.failedAt = new Date().getTime();
        task.duration = (task.failedAt - task.startedAt) / 1000;

        // アクティブタスクから失敗タスクに移動
        delete queue.activeTasks[taskId];
        queue.failedTasks.push(task);
        queue.lastUpdated = new Date().getTime();

        this.saveQueue(queue);

        log('ERROR', `タスク最終失敗: ${taskId}, 再試行回数=${task.retryCount}, エラー=${errorMessage}`);

        return {
          status: 'failed',
          taskId: taskId,
          retryCount: task.retryCount,
          errorMessage: errorMessage,
          duration: task.duration
        };
      }

    } catch (error) {
      log('ERROR', 'タスク失敗処理エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * タスクのステータスを更新
   * @param {string} taskId - タスクID
   * @param {Object} updateData - 更新データ
   * @return {Object} 更新結果
   */
  updateTaskStatus(taskId, updateData) {
    try {
      const queue = this.getQueue();
      if (!queue) {
        return { status: 'error', message: 'キューが見つかりません' };
      }

      // アクティブタスクから検索
      let task = queue.activeTasks[taskId];
      let taskLocation = 'active';

      // アクティブタスクで見つからない場合は通常のキューを検索
      if (!task) {
        task = queue.tasks.find(t => t.taskId === taskId);
        taskLocation = 'queued';
      }

      if (!task) {
        return { status: 'error', message: 'タスクが見つかりません' };
      }

      // 更新可能なフィールドを更新
      const updatableFields = ['progress', 'currentStep', 'customData', 'estimatedDuration'];
      updatableFields.forEach(field => {
        if (updateData.hasOwnProperty(field)) {
          task[field] = updateData[field];
        }
      });

      task.updatedAt = new Date().getTime();
      queue.lastUpdated = new Date().getTime();

      this.saveQueue(queue);

      log('DEBUG', `タスクステータス更新: ${taskId}, 場所=${taskLocation}`);

      return {
        status: 'updated',
        taskId: taskId,
        taskLocation: taskLocation,
        updatedFields: Object.keys(updateData)
      };

    } catch (error) {
      log('ERROR', 'タスクステータス更新エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * キューの統計情報を取得
   * @return {Object} 統計情報
   */
  getQueueStatistics() {
    try {
      const queue = this.getQueue();
      if (!queue) {
        return this.getEmptyStatistics();
      }

      const stats = {
        totalTasks: queue.tasks.length + Object.keys(queue.activeTasks).length + 
                   queue.completedTasks.length + queue.failedTasks.length,
        queuedTasks: queue.tasks.length,
        activeTasks: Object.keys(queue.activeTasks).length,
        completedTasks: queue.completedTasks.length,
        failedTasks: queue.failedTasks.length,
        concurrencyUtilization: Math.round((Object.keys(queue.activeTasks).length / this.maxConcurrentTasks) * 100),
        averageWaitTime: 0,
        averageProcessingTime: 0,
        successRate: 0,
        priorityDistribution: {},
        typeDistribution: {},
        lastUpdated: queue.lastUpdated
      };

      // 平均待機時間の計算
      if (queue.completedTasks.length > 0) {
        const totalWaitTime = queue.completedTasks.reduce((sum, task) => {
          return sum + ((task.startedAt || task.createdAt) - task.createdAt);
        }, 0);
        stats.averageWaitTime = Math.round(totalWaitTime / queue.completedTasks.length / 1000);
      }

      // 平均処理時間の計算
      if (queue.completedTasks.length > 0) {
        const totalProcessingTime = queue.completedTasks.reduce((sum, task) => {
          return sum + (task.duration || 0);
        }, 0);
        stats.averageProcessingTime = Math.round(totalProcessingTime / queue.completedTasks.length);
      }

      // 成功率の計算
      const totalProcessedTasks = queue.completedTasks.length + queue.failedTasks.length;
      if (totalProcessedTasks > 0) {
        stats.successRate = Math.round((queue.completedTasks.length / totalProcessedTasks) * 100);
      }

      // 優先度分布の計算
      const allTasks = [
        ...queue.tasks,
        ...Object.values(queue.activeTasks),
        ...queue.completedTasks,
        ...queue.failedTasks
      ];

      allTasks.forEach(task => {
        // 優先度分布
        const priority = task.priority || 5;
        stats.priorityDistribution[priority] = (stats.priorityDistribution[priority] || 0) + 1;
        
        // タイプ分布
        const type = task.type || 'unknown';
        stats.typeDistribution[type] = (stats.typeDistribution[type] || 0) + 1;
      });

      return stats;

    } catch (error) {
      log('ERROR', 'キュー統計取得エラー', error);
      return this.getEmptyStatistics();
    }
  }

  /**
   * 特定タスクの詳細情報を取得
   * @param {string} taskId - タスクID
   * @return {Object|null} タスク詳細情報
   */
  getTaskDetails(taskId) {
    try {
      const queue = this.getQueue();
      if (!queue) {
        return null;
      }

      // 全てのタスクコレクションから検索
      const allTasks = [
        ...queue.tasks.map(t => ({...t, location: 'queued'})),
        ...Object.values(queue.activeTasks).map(t => ({...t, location: 'active'})),
        ...queue.completedTasks.map(t => ({...t, location: 'completed'})),
        ...queue.failedTasks.map(t => ({...t, location: 'failed'}))
      ];

      const task = allTasks.find(t => t.taskId === taskId);
      if (!task) {
        return null;
      }

      // 待機時間と処理時間を計算
      const currentTime = new Date().getTime();
      const waitTime = task.startedAt ? 
        Math.round((task.startedAt - task.createdAt) / 1000) : 
        Math.round((currentTime - task.createdAt) / 1000);

      const processingTime = task.location === 'active' ? 
        Math.round((currentTime - task.startedAt) / 1000) :
        (task.duration || 0);

      return {
        ...task,
        waitTime: waitTime,
        processingTime: processingTime,
        queuePosition: task.location === 'queued' ? 
          queue.tasks.findIndex(t => t.taskId === taskId) + 1 : null
      };

    } catch (error) {
      log('ERROR', 'タスク詳細取得エラー', error);
      return null;
    }
  }

  /**
   * キュー全体の情報を取得
   * @param {Object} options - 取得オプション
   * @return {Object} キュー情報
   */
  getQueueInfo(options = {}) {
    try {
      const queue = this.getQueue();
      if (!queue) {
        return { status: 'empty', message: 'キューが見つかりません' };
      }

      const result = {
        status: 'active',
        statistics: this.getQueueStatistics(),
        queue: {
          tasks: queue.tasks.map(this.formatTaskForDisplay),
          activeTasks: Object.values(queue.activeTasks).map(this.formatTaskForDisplay),
          recentCompleted: queue.completedTasks.slice(-10).map(this.formatTaskForDisplay),
          recentFailed: queue.failedTasks.slice(-10).map(this.formatTaskForDisplay)
        },
        lastUpdated: queue.lastUpdated
      };

      // 詳細情報を含める場合
      if (options.includeDetails) {
        result.queue.allCompleted = queue.completedTasks.map(this.formatTaskForDisplay);
        result.queue.allFailed = queue.failedTasks.map(this.formatTaskForDisplay);
      }

      return result;

    } catch (error) {
      log('ERROR', 'キュー情報取得エラー', error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * 古いタスクをクリーンアップ
   * @param {number} maxAgeHours - 最大保持時間（時間）
   * @return {Object} クリーンアップ結果
   */
  cleanupOldTasks(maxAgeHours = 24) {
    try {
      const queue = this.getQueue();
      if (!queue) {
        return { status: 'no_queue', cleaned: 0 };
      }

      const cutoffTime = new Date().getTime() - (maxAgeHours * 60 * 60 * 1000);
      let cleanedCount = 0;

      // 完了タスクのクリーンアップ
      const originalCompletedCount = queue.completedTasks.length;
      queue.completedTasks = queue.completedTasks.filter(task => {
        const keep = (task.completedAt || task.updatedAt) > cutoffTime;
        if (!keep) cleanedCount++;
        return keep;
      });

      // 失敗タスクのクリーンアップ
      const originalFailedCount = queue.failedTasks.length;
      queue.failedTasks = queue.failedTasks.filter(task => {
        const keep = (task.failedAt || task.updatedAt) > cutoffTime;
        if (!keep) cleanedCount++;
        return keep;
      });

      queue.lastUpdated = new Date().getTime();
      this.saveQueue(queue);

      log('INFO', `古いタスクをクリーンアップ: ${cleanedCount}件削除 (完了: ${originalCompletedCount - queue.completedTasks.length}, 失敗: ${originalFailedCount - queue.failedTasks.length})`);

      return {
        status: 'cleaned',
        cleaned: cleanedCount,
        completedCleaned: originalCompletedCount - queue.completedTasks.length,
        failedCleaned: originalFailedCount - queue.failedTasks.length
      };

    } catch (error) {
      log('ERROR', 'タスククリーンアップエラー', error);
      return { status: 'error', message: error.message };
    }
  }

  // === プライベートメソッド ===

  /**
   * キューをキャッシュから取得
   * @return {Object|null} キューデータ
   */
  getQueue() {
    const cached = this.cache.get(this.queueKey);
    return cached ? JSON.parse(cached) : null;
  }

  /**
   * キューをキャッシュに保存
   * @param {Object} queue - キューデータ
   */
  saveQueue(queue) {
    this.cache.put(this.queueKey, JSON.stringify(queue), 21600); // 6時間有効
  }

  /**
   * タスクIDを生成
   * @return {string} 生成されたタスクID
   */
  generateTaskId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000);
    return `TASK_${timestamp}_${random}`;
  }

  /**
   * 空の統計情報を返す
   * @return {Object} 空の統計情報
   */
  getEmptyStatistics() {
    return {
      totalTasks: 0,
      queuedTasks: 0,
      activeTasks: 0,
      completedTasks: 0,
      failedTasks: 0,
      concurrencyUtilization: 0,
      averageWaitTime: 0,
      averageProcessingTime: 0,
      successRate: 0,
      priorityDistribution: {},
      typeDistribution: {},
      lastUpdated: 0
    };
  }

  /**
   * 表示用にタスクをフォーマット
   * @param {Object} task - タスクオブジェクト
   * @return {Object} フォーマットされたタスク
   */
  formatTaskForDisplay(task) {
    return {
      taskId: task.taskId,
      type: task.type,
      status: task.status,
      priority: task.priority,
      retryCount: task.retryCount,
      progress: task.progress,
      currentStep: task.currentStep,
      createdAt: task.createdAt,
      startedAt: task.startedAt,
      updatedAt: task.updatedAt,
      duration: task.duration,
      errorMessage: task.errorMessage
    };
  }
}