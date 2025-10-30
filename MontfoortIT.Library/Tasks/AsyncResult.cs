namespace MontfoortIT.Library.Tasks
{
    public class AsyncResult<S>
    {
        public bool Finished { get; private set; }

        public S Result { get; private set; }

        internal AsyncResult()
        {
            Finished = false;
        }

        internal AsyncResult(S result)
        {
            Finished = true;
            Result = result;
        }

        public AsyncResult(bool finished)
        {
            Finished = finished;
        }
    }
}