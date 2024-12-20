using AsnReport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace preStockReportService{
    class CReports
    {
        private COracle m_oracle;
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        
        public void getPrestocktakeReport(System.Diagnostics.EventLog system_events)
        {
            try{
                //SEMPROD es el SID de SEM Producción
                m_oracle = new COracle("192.168.0.23", "SEMPROD");
                excel m_excel = new excel();
                //String pathReporterror = "";
                CUtils utils = new CUtils();
                //String error = "";
                String message;

                message = m_oracle.preStocktake(system_events) ? "Se generó el Reporte de PreStocktake. " : "NO se generó el Reporte de PreStocktake. " ;
                system_events.WriteEntry(message);
            }
            catch (Exception ex) {
                system_events.WriteEntry("getPrestocktakeReport - Ocurrió un error al Construir Reporte. \n" + ex.Message);
            }
        }
    }
}
