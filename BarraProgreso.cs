using System.Text;

namespace ChequeBancarioOpenXML
{
    public class BarraProgreso : IDisposable, IProgress<(double progreso, string operacion)>
    {
        private const int BLOCK_COUNT = 30;
        private readonly TimeSpan _animationInterval = TimeSpan.FromSeconds(1.0 / 8);
        private const string ANIMATION = @"|/-\";

        private readonly Timer _timer;
        private double _currentProgress = 0;
        private string _currentText = string.Empty;
        private string _currentOperation = "Procesando";
        private bool _disposed = false;
        private int _animationIndex = 0;

        public BarraProgreso()
        {
            _timer = new Timer(TimerHandler!);

            if (!Console.IsOutputRedirected)
            {
                ResetTimer();
            }
        }

        public void Report((double progreso, string operacion) value)
        {
            Interlocked.Exchange(ref _currentProgress, value.progreso);
            _currentOperation = value.operacion ?? _currentOperation;
        }

        public void Report(double progreso, string operacion)
        {
            Report((progreso, operacion));
        }

        private void TimerHandler(object state)
        {
            lock (_timer)
            {
                if (_disposed) return;

                int progressBlockCount = (int)(_currentProgress * BLOCK_COUNT);
                int percent = (int)(_currentProgress * 100);
                string text = string.Format("[{0}{1}] {2,3}% {3} {4}",
                    new string('█', progressBlockCount),
                    new string('░', BLOCK_COUNT - progressBlockCount),
                    percent,
                    ANIMATION[_animationIndex++ % ANIMATION.Length],
                    _currentOperation);

                UpdateText(text);
                ResetTimer();
            }
        }

        private void UpdateText(string text)
        {
            int commonPrefixLength = 0;
            int commonLength = Math.Min(_currentText.Length, text.Length);

            while (commonPrefixLength < commonLength &&
                   text[commonPrefixLength] == _currentText[commonPrefixLength])
            {
                commonPrefixLength++;
            }

            StringBuilder outputBuilder = new StringBuilder();
            outputBuilder.Append('\b', _currentText.Length - commonPrefixLength);
            outputBuilder.Append(text.Substring(commonPrefixLength));

            int overlapCount = _currentText.Length - text.Length;
            if (overlapCount > 0)
            {
                outputBuilder.Append(' ', overlapCount);
                outputBuilder.Append('\b', overlapCount);
            }

            Console.Write(outputBuilder);
            _currentText = text;
        }

        private void ResetTimer()
        {
            _timer.Change(_animationInterval, TimeSpan.FromMilliseconds(-1));
        }

        public void Dispose()
        {
            lock (_timer)
            {
                _disposed = true;
                // Mostrar el 100% completado antes de limpiar
                UpdateText($"[{new string('█', BLOCK_COUNT)}] 100% - {_currentOperation}");
                Console.WriteLine();
                _timer.Dispose();
            }
        }
    }
}