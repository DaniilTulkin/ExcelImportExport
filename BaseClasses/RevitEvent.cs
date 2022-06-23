using Autodesk.Revit.UI;
using System;

namespace ExcelImportExport
{
    public class RevitEvent : IExternalEventHandler
    {
        private Action<UIApplication> _action;
        private readonly ExternalEvent _externalEvent;
        public RevitEvent()
        {
            _externalEvent = ExternalEvent.Create(this);
        }
        public void Run(Action<UIApplication> action)
        {
            _action = action;
            _externalEvent.Raise();
        }
        public void Execute(UIApplication app)
        {
            try
            {
                _action?.Invoke(app);
            }
            catch (Exception ex)
            {
            }
        }
        public string GetName() => nameof(RevitEvent);
    }
}
