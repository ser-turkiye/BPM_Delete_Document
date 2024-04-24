package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.metaDataComponents.IArchiveClass;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;

import static ser.Utils.loadTableRows;

public class CancelProcess extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    private IDocument mainDocument;
    JSONObject projects = new JSONObject();
    private ITask mainTask;
    private ITask task;
    private String notes;
    String prjn = "";
    List<String> docs = new ArrayList<>();
    IDocument mailTemplate = null;
    @Override
    protected Object execute() {
        helper = new ProcessHelper(getSes());

        task = getEventTask();
        mainTask = getEventTask().findParentTask();
        IProcessInstance processInstance = mainTask.getProcessInstance();
        String currentUser = getSes().getUser().getFullName();
        IUser processOwner = processInstance.getOwner();
        String uniqueId = UUID.randomUUID().toString();

        if (mainTask == null) return resultError("OBJECT CLIENT ID is NULL or not of Type ITask");
        try {
            com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

            Utils.session = getSes();
            Utils.bpm = getBpm();
            Utils.server = Utils.session.getDocumentServer();
            Utils.loadDirectory(Conf.CancelProcess.MainPath);
            this.helper = new ProcessHelper(Utils.session);
            //helper = new ProcessHelper(getSes());
            log.info("----DeleteDocument Process Agent Started -----:" + mainTask.getID());


            int cnt = 0;
            JSONObject dbks = new JSONObject();
            this.helper = new ProcessHelper(Utils.session);
            JSONObject mcfg = Utils.getMailConfig();
            dbks.put("DoxisLink", mcfg.getString("webBase") + helper.getTaskURL(processInstance.getID()));

            String prjn = "",  mdno = "", mdrn = "", mdnm = "";
            IInformationObjectLinks links = mainTask.getProcessInstance().getLoadedInformationObjectLinks();
            for (ILink link : links.getLinks()) {
                IInformationObject xdoc = link.getTargetInformationObject();
                String taskName = xdoc.getDescriptorValue("ccmPrjDocWFTaskName");
                if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)) {
                    continue;
                }
                prjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);

                if(xdoc != null &&  Utils.hasDescriptor((IInformationObject) xdoc, Conf.Descriptors.ProjectNo)){
                    prjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                }
                if(xdoc != null &&  Utils.hasDescriptor((IInformationObject) xdoc, Conf.Descriptors.DocNumber)){
                    mdno = xdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(xdoc != null &&  Utils.hasDescriptor((IInformationObject) xdoc, Conf.Descriptors.Revision)){
                    mdrn = xdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(xdoc != null &&  Utils.hasDescriptor((IInformationObject) xdoc, Conf.Descriptors.Name)){
                    mdnm = xdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }
                cnt++;

                Date tbgn = null, tend = new Date();
                if(task.getReadyDate() != null){
                    tbgn = task.getReadyDate();
                }
                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }
                String rcvf = "", rcvo = "";
                if(mainTask.getPreviousWorkbasket() != null){
                    rcvf = mainTask.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                dbks.put("DocNo" + cnt, (mdno != null  ? mdno : ""));
                dbks.put("RevNo" + cnt, (mdrn != null  ? mdrn : ""));
                dbks.put("Title" + cnt, xdoc.getDisplayName());
                dbks.put("Task" + cnt, mainTask.getName());
                dbks.put("DocName" + cnt, (mdnm != null  ? mdnm : ""));
                dbks.put("CancelledOn" + cnt, (rcvo != null ? rcvo : ""));
                dbks.put("ProcessTitle" + cnt, (processInstance != null ? processInstance.getDisplayName() : ""));
                dbks.put("ProjectNo" + cnt, (prjn != null  ? prjn : ""));
                dbks.put("DoxisDocLink" + cnt, mcfg.get("webBase") + helper.getDocumentURL(xdoc.getID()));
                dbks.put("DoxisDocLink" + cnt + ".Text", "( Link )");
                //dbks.put("DoxisLink" + (cnt > 9 ? cnt : "0" + cnt), mcfg.getString("webBase") + helper.getTaskURL(mainTask.getID()));
                docs.add(xdoc.getDescriptorValue(Conf.Descriptors.DocNumber));
            }

            Long prevTaskID = this.getEventTask().getPreviousTaskNumericID();
            ITask prevTask = this.getEventTask().getProcessInstance().findTaskByNumericID(this.getEventTask().getPreviousTaskNumericID());
            log.info("Previev task name :" + (prevTask != null ? prevTask.getName() : "---"));

            String mtpn = "PROCESS_CANCEL_MAIL";
            projects = Utils.getProjectWorkspaces(this.helper);
            mailTemplate = null;
            for(String prjnmbr : projects.keySet()){
                IInformationObject prjt = (IInformationObject) projects.get(prjnmbr);
                IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.CancelProcess);
                if(dtpl == null){continue;}
                mailTemplate = dtpl;
                break;
            }
            dbks.put("docs", String.join(", ", docs));
            if (mailTemplate == null) {
                log.info("Template-Document [ " + mtpn + " ] not found.");
                //throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            } else {
                String tplMailPath = Utils.exportDocument(mailTemplate, Conf.CancelProcess.MainPath, mtpn + "[" + uniqueId + "]");

                loadTableRows(tplMailPath, 0, "Task", 0, docs.size());
                String mailExcelPath = Utils.saveToExcel(tplMailPath, 0,
                        Conf.CancelProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
                );
                String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.CancelProcessSheetIndex.Mail,Conf.CancelProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

                String umail = processOwner.getEMailAddress();
                List<String> mails = new ArrayList<>();

                if (umail != null) {
                    mails.add(umail);
                    JSONObject mail = new JSONObject();
                    mail.put("To", String.join(";", mails));
                    mail.put("Subject", "Process Cancelled -" + (prevTask != null ? prevTask.getName() : "") + " - " + (mdno != null  ? mdno : "") + "/" + (mdrn != null  ? mdrn : ""));
                    mail.put("BodyHTMLFile", mailHtmlPath);
                    Utils.sendHTMLMail(mail, null);
                } else {
                    log.info("Mail adress is null :" + processOwner.getFullName());
                }
                log.info("----DeleteDocument Process Agent Finished -----");
            }
        } catch (Exception e) {
            log.error("Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        return resultSuccess("Agent Finished Succesfully");
    }
}
