using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Ctse.Outlook.Toolbox
{
    public static class ComSafeHelper
    {
        public static ComSafe<T> For<T>(T item)
        {
            return new ComSafe<T>(item);
        }

        public static void ReleaseItems<T>(IEnumerable<T> items)
        {
            foreach (T obj in items)
                ComSafeHelper.Release((object)obj);
        }

        public static void Release(object item)
        {
            if (item == null)
                return;
            try
            {
                Marshal.ReleaseComObject(item);
            }
            catch (Exception ex)
            {
            }
        }
    }

    public class ComSafe<T> : IDisposable
    {
        private readonly object _disposeLock = new object();
        private T _instance;
        private readonly bool _shouldRelease;
        private readonly IDisposable _owner;
        private bool _isDisposed;

        internal ComSafe(T instance)
          : this(instance, false, (IDisposable)null)
        {
        }

        private ComSafe(T instance, bool shouldRelease, IDisposable owner)
        {
            this._instance = instance;
            this._shouldRelease = shouldRelease;
            this._owner = owner;
        }

        public ComSafe<TRet> With<TRet>(Func<T, TRet> accessor)
        {
            try
            {
                return new ComSafe<TRet>(accessor(this._instance), true, (IDisposable)this);
            }
            catch (Exception ex)
            {
                this.Dispose();
                throw;
            }
        }

        public TRet Do<TRet>(Func<T, TRet> action)
        {
            try
            {
                return action(this._instance);
            }
            finally
            {
                this.Dispose();
            }
        }

        public void Do(Action<T> action)
        {
            try
            {
                action(this._instance);
            }
            finally
            {
                this.Dispose();
            }
        }

        public void Dispose()
        {
            if (this._isDisposed)
                return;
            lock (this._disposeLock)
            {
                if (this._isDisposed)
                    return;
                if ((object)this._instance != null)
                {
                    if (this._shouldRelease)
                        Marshal.ReleaseComObject((object)this._instance);
                    this._instance = default(T);
                }
                if (this._owner != null)
                    this._owner.Dispose();
                this._isDisposed = true;
            }
        }
    }
}
