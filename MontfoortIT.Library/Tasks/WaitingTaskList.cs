using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MontfoortIT.Library.Tasks
{
    public class WaitingTaskList
    {
        private int _maxConcurrentTasks;
        
        private System.Collections.Concurrent.ConcurrentDictionary<int, Task> _concurrentTasks = new System.Collections.Concurrent.ConcurrentDictionary<int, Task>();

        public WaitingTaskList(int maxConcurrentTasks)
        {
            _maxConcurrentTasks = maxConcurrentTasks;
        }
        
        public async Task<AsyncResult<S>> AddAsync<S>(Task task, int waitTimeoutInMiliseconds = int.MaxValue, CancellationToken cancelationToken = default(CancellationToken))
        {
            if (task == null)
                return new AsyncResult<S>();

            // add it to the task list, do not wait for it
#pragma warning disable CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed
            _concurrentTasks.AddOrUpdate(task.Id, task, (i, t) => t);
#pragma warning restore CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed

            if (!CanAdd())
                return await WhenAnyAsync<S>(waitTimeoutInMiliseconds, cancelationToken);

            return new AsyncResult<S>();
        }

        public bool CanAdd()
        {
            return _concurrentTasks.Count < _maxConcurrentTasks;
        }

        public void CleanExceptions()
        {
            var faultedTasks = _concurrentTasks.Select(t => t.Value).Where(c => c.Status == TaskStatus.Faulted).ToArray();
            foreach (var faultedTask in faultedTasks)
            {
                Task dummy;
                _concurrentTasks.TryRemove(faultedTask.Id, out dummy);
            }
        }

        public Task AddAsync(Task task, int waitTimeout = int.MaxValue)
        {            
            return AddAsync<int>(task, waitTimeout);
        }


        public async Task<bool> WaitForActivePositionAsync(int timeoutInMilliseconds)
        {
            if(_concurrentTasks.Count >= _maxConcurrentTasks)
            {
                Task whenAny = WhenAnyAsync();
                var delayTask = Task.Delay(timeoutInMilliseconds);

                Task finishedTask = await Task.WhenAny(whenAny, delayTask);
                return _concurrentTasks.Count < _maxConcurrentTasks;                
            }
            return true;
        }

        public async Task WhenAllFinishedAsync()
        {
            var completed = _concurrentTasks.Where(s => s.Value.IsCompleted);
            foreach (var c in completed)
            {
                if (c.Value.IsFaulted)
                    throw c.Value.Exception;

                _concurrentTasks.TryRemove(c.Key, out var v);
            }

            if (_concurrentTasks.Count > 0)
                await Task.WhenAll(_concurrentTasks.Select(c=>c.Value).ToArray());
        }

        public async Task WhenAllFinishedAsync(CancellationToken token)
        {
            var completed = _concurrentTasks.Where(s => s.Value.IsCompleted);
            foreach (var c in completed)
            {
                _concurrentTasks.TryRemove(c.Key, out var v);
            }

            if (_concurrentTasks.Count > 0)
            {
                Task cancelTask = Task.Delay(int.MaxValue, token);
                var allTask = Task.WhenAll(_concurrentTasks.Select(c => c.Value).ToArray());
                await Task.WhenAny(allTask, cancelTask);
            }
        }

        public async Task WhenAllFinishedAsync(int timeout)
        {
            if (_concurrentTasks.Count > 0)
            {
                Task whenAll = WhenAllFinishedAsync();
                var delayTask = Task.Delay(timeout);

                await Task.WhenAny(whenAll, delayTask);
            }
        }

        public async Task WhenAllFinishedAsync(int timeout, CancellationToken cancellationToken)
        {
            if (_concurrentTasks.Count > 0)
            {
                Task whenAll = WhenAllFinishedAsync();
                var delayTask = Task.Delay(timeout, cancellationToken);

                await Task.WhenAny(whenAll, delayTask);
            }
        }


        private async Task<bool> WhenAnyAsync()
        {
            if (_concurrentTasks.Count > 0)
            {
                Task dummy;
                try
                {
                    var task = await Task.WhenAny(_concurrentTasks.Select(c => c.Value).ToArray());
                    _concurrentTasks.TryRemove(task.Id, out dummy);

                    return true;
                }
                catch (Exception e)
                {
                    var faultedTask = _concurrentTasks.Select(t => t.Value).Where(c => c.Status == TaskStatus.Faulted && c.Exception == e).FirstOrDefault();

                    _concurrentTasks.TryRemove(faultedTask.Id, out dummy);

                    throw;
                }
            }
            return false;
        }

        public async Task<AsyncResult<S>> WhenAnyAsync<S>(int waitTimeoutInMiliseconds = int.MaxValue, CancellationToken cancellationToken = default(CancellationToken))
        {
            if (_concurrentTasks.Count > 0)
            {
                Task dummy;
                try
                {
                    Task task = _concurrentTasks.Select(c=>c.Value).FirstOrDefault(k=> k.IsCompleted);
                    if (task == null)
                    {
                        bool cancellationSet = cancellationToken != default(CancellationToken);
                        if (waitTimeoutInMiliseconds < int.MaxValue || cancellationSet)
                        {
                            Task delay;
                            if(cancellationSet)
                                delay = Task.Delay(waitTimeoutInMiliseconds, cancellationToken);
                            else
                                delay = Task.Delay(waitTimeoutInMiliseconds);

                            List<Task> waitList = new List<Task>(_concurrentTasks.Select(c => c.Value).ToArray())
                            {
                                delay
                            };
                            task = await Task.WhenAny(waitList);
                        }
                        else
                            task = await Task.WhenAny(_concurrentTasks.Select(c => c.Value).ToArray());
                        
                    }

                    bool done = _concurrentTasks.TryRemove(task.Id, out dummy);

                    if (task is Task<S> sTask)
                    {
                        S result = await sTask;
                        return new AsyncResult<S>(result);
                    }
                    else
                    {
                        await task;
                        return new AsyncResult<S>(true);
                    }
                }                
                catch (Exception e)
                {
                    var faultedTask = _concurrentTasks.Select(t => t.Value).Where(c => c.Status == TaskStatus.Faulted && c.Exception == e).FirstOrDefault();
                             
                    if(faultedTask!=null)           
                        _concurrentTasks.TryRemove(faultedTask.Id, out dummy);

                    throw;
                }
            }
            return new AsyncResult<S>();
        }
    }
}
