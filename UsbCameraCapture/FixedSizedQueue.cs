using System;
using System.Collections.Concurrent;

namespace UsbCameraCapture
{
    public class FixedSizedQueue<T>
    {
        private ConcurrentQueue<T> _q;
        private object _lockObject = new object();

        public FixedSizedQueue(int limit = 30)
        {
            _q = new ConcurrentQueue<T>();
            Limit = limit;
        }

        public int Limit { get; set; }

        public void Enqueue(T obj)
        {
            _q.Enqueue(obj);
            lock (_lockObject)
            {
                T overflow;
                while (_q.Count > Limit && _q.TryDequeue(out overflow)) ;
            }
        }

        public T Dequeue()
        {
            T result;
            var ret = _q.TryDequeue(out result);
            if (ret)
            {
                return result;
            }
            else 
            {
                throw new InvalidOperationException();
            }
        }

        public void Clear()
        {
            lock (_lockObject)
            {
                T overflow;
                while (_q.Count > 0 && _q.TryDequeue(out overflow)) ;
            }
        }
    }
}
