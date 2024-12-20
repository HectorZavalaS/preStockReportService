using smtLocations.Class;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsnReport{
    class CMailSender
    {
        private COracle m_oracle;
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public void sendMail(System.Diagnostics.EventLog system_events)
        {
            try  {
                //MXPRD es MX Producción
                m_oracle = new COracle("172.25.0.15", "MXPRD");
                excel m_excel = new excel();
                String pathReport = "";
                String pathReporterror = "";
                CUtils utils = new CUtils();
                String error = "";
                //system_events.WriteEntry("Obteniendo registros de base de datos de Oracle.");
                if(m_oracle.get_ASN_report(ref pathReport, system_events)) { 

                    List<string> lstArchivos = new List<string>();
                    lstArchivos.Add(pathReport);

                    if(m_oracle.get_error_report(ref pathReporterror, system_events))
                        lstArchivos.Add(pathReporterror);
                    String mails = "asn-sem@siix-sem.com.mx;warehouse.receiving@SIIX-SEM.com.mx;ruben.regis@SIIX-SEM.com.mx;kenny.manzanilla@SIIX-SEM.com.mx;christian.gonzalez@siix-sem.com.mx;nicolas.delangel@siix-sem.com.mx;cristobal.munoz@siix-sem.com.mx;antonio.hernandez@siix-sem.com.mx;javier.gallardo@siix-sem.com.mx;victor.moreno@siix-sem.com.mx;minerva.gaitan@siix-sem.com.mx;dulce.loredo@siix-sem.com.mx;raymundo.salas@siix-sem.com.mx;luis.torres@siix.mx";
                    //String mails = "antonio.hernandez@siix-sem.com.mx";

                    //creamos nuestro objeto de la clase que hicimos
                    CMail oMail = new CMail("ASN_Report@siix.mx", mails,
                                         "ASNs Report", "ASNs Report", lstArchivos);

                    oMail.Message = "Se anexa reporte de ASNs / Attached you will find ASNs report.<br><br> Saludos / Regards.";

                    //y enviamos
                    if (oMail.enviaMail(ref error))  {
                        system_events.WriteEntry("Se envió por E-mail ASNs Report.");

                    }  else  {
                        system_events.WriteEntry("No se envió el mail: " + oMail.error + "  \n" + error);
                       //logger.Error("No se envio el mail: " + oMail.error);

                    }
                }
            }  catch(Exception ex)  {
                system_events.WriteEntry("Ocurrió un error al Construir Reporte. " + ex.Message);
            }
        }
    }
}
