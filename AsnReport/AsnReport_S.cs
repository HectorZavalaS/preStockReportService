using preStockReportService;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;


namespace AsnReport
{
    public partial class AsnReport_S : ServiceBase
    {
        CReports m_report;
        Timer timer = new Timer();
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();


        public AsnReport_S()
        {
            InitializeComponent();
            m_report = new CReports();
            system_events = new System.Diagnostics.EventLog();

            if (!EventLog.SourceExists("PreStockTake Report")) {
                EventLog.CreateEventSource("PreStockTake Report", "Application");
            }
            system_events.Source = "PreStockTake Report";
            system_events.Log = "Application";

        }

        protected override void OnStart(string[] args = null)
        {
            AsnReport_t1.Start();
            /*
            try {
                system_events.WriteEntry("Iniciado servicio de reporte PreStockTake. ");
                timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
                timer.Interval = 7200000; //number in milisecinds  (2 HOURS)
                timer.Enabled = true;
                m_report.getPrestocktakeReport(system_events);
            }  catch (Exception ex)  {
                system_events.WriteEntry("Ocurrio un error al iniciar el Timer. " + ex.Message);
                //logger.Error(ex, "Ocurrio un error al iniciar el Timer.");
            }*/
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            try
            {
                system_events.WriteEntry("Iniciado servicio de reporte PreStockTake. ");
                m_report.getPrestocktakeReport(system_events);
            } catch (Exception ex)
            {
                system_events.WriteEntry("Ocurrio un error al iniciar el Timer. " + ex.Message);
            }
        }

        protected override void OnStop() {
            AsnReport_t1.Stop();
        }

        private void AsnReport_t_Tick(object sender, EventArgs e)  {
            
        }

    }
}
