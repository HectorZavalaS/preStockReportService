
namespace AsnReport
{
    partial class AsnReport_S
    {
        /// <summary> 
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary> 
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.system_events = new System.Diagnostics.EventLog();
            this.AsnReport_t1 = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.system_events)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.AsnReport_t1)).BeginInit();
            // 
            // AsnReport_t1
            // 
            this.AsnReport_t1.Enabled = true;
            this.AsnReport_t1.Interval = 3600000D;//7200000 3600000
            this.AsnReport_t1.Elapsed += new System.Timers.ElapsedEventHandler(this.OnElapsedTime);
            // 
            // AsnReport_S
            // 
            this.ServiceName = "PreStocktake Report";
            ((System.ComponentModel.ISupportInitialize)(this.system_events)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.AsnReport_t1)).EndInit();

        }

        #endregion
        private System.Diagnostics.EventLog system_events;
        private System.Timers.Timer AsnReport_t1;
    }
}
