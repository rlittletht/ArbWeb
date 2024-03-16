using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Application = System.Windows.Forms.Application;

namespace ArbWeb
{
    public class TaskYielder
    {
        public delegate void Fun();

        class ProgressCounter
        {
            readonly CancellationTokenSource m_tokenSource = new CancellationTokenSource();
            public CancellationToken Token { get; set; }

            public IAppContext Context { get; set; }

            public ProgressCounter(IAppContext appContext)
            {
                Context = appContext;
                Token = m_tokenSource.Token;
            }

            public void Cancel()
            {
                m_tokenSource.Cancel();
            }
        }

        ProgressCounter StartCounter(IAppContext context, string sProgressName)
        {
            ProgressCounter counter = new ProgressCounter(context);

            Task task = new Task(
                () =>
                {
                    int count = 0;

                    while (!counter.Token.IsCancellationRequested)
                    {
                        counter.Context.StatusReport.AddMessage($"{sProgressName}: {count++}");
                        Thread.Sleep(1000);
                    }
                },
                counter.Token);

            task.Start();

            return counter;
        }

        public static void RunTaskWithYieldingWait(
            IAppContext context,
            Fun fun,
            bool ShowCounter = false,
            string sProgressName = "")
        {
            TaskYielder yielder = new TaskYielder(context, fun, ShowCounter, sProgressName);
        }

        public TaskYielder(IAppContext context, Fun fun, bool ShowCounter = false, string sProgressName = "")
        {
            ProgressCounter counter = null;

            if (ShowCounter)
            {
                counter = StartCounter(context, sProgressName);
            }

            Task task = new Task(() => { fun(); });

            task.Start();

            while (!task.IsCompleted)
            {
                task.Wait(500);
                Application.DoEvents();
            }

            counter?.Cancel();
        }
    }
}
