package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

public class DeletionDocs extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    JSONObject projects = new JSONObject();
    private ProcessHelper helper;
    private IDocument mainDocument;
    private ITask task;
    List<String> docs = new ArrayList<>();
    private String notes;
    @Override
    protected Object execute() {
        if (getEventTask() == null)
            return resultError("Null Document object");

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.Paths.MainPath);

        task = getEventTask();
        processInstance = task.getProcessInstance();
        String uniqueId = UUID.randomUUID().toString();

        List<String> mails = new ArrayList<>();
        log.info("Mail To : " + String.join(";", mails));

        IRole role = getDocumentServer().getRoleByName(getSes(), Conf.RoleNames.Admins);
        if(role != null) {
            IUser[] usrs = role.getUserMembers();
            for (IUser usr : usrs) {
                String mail = usr.getEMailAddress();
                mail = (mail == null ? "" : mail);
                if(mail.isEmpty()){continue;}
                if(mails.contains(mail)){continue;}
                mails.add(mail);
            }
        }


        if(mails.size() == 0){
            return resultSuccess("Agent passed. No mail address.");
        }

        try {
            com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);
            this.helper = new ProcessHelper(getSes());
            //task.commit();

            String mtpn = "DOCUMENT_DELETION_MAIL";
            projects = Utils.getProjectWorkspaces(helper);
            IDocument mtpl = null;
            for(String prjn : projects.keySet()){
                IInformationObject prjt = (IInformationObject) projects.get(prjn);
                IDocument dtpl = Utils.getTemplateDocument(prjt, mtpn);
                if(dtpl == null){continue;}
                mtpl = dtpl;
                break;
            }
            if(mtpl == null){
                log.info("Template-Document [ " + mtpn + " ] not found.");
                //throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }else {
                JSONObject dbks = new JSONObject();
                JSONObject mcfg = Utils.getMailConfig();
                dbks.put("DoxisLink", mcfg.getString("webBase") + helper.getTaskURL(processInstance.getID()));

                int cnt = 0;
                this.helper = new ProcessHelper(Utils.session);
                String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                IInformationObjectLinks links = task.getProcessInstance().getLoadedInformationObjectLinks();
                for (ILink link : links.getLinks()) {
                    IInformationObject xdoc = link.getTargetInformationObject();
                    String taskName = xdoc.getDescriptorValue("ccmPrjDocWFTaskName");
                    if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)) {
                        continue;
                    }
                    prjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    //this.deleteDocument(xdoc);
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
                    docs.add(xdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class));

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
                    if(task.getPreviousWorkbasket() != null){
                        rcvf = task.getPreviousWorkbasket().getFullName();
                    }
                    if(tbgn != null){
                        rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                    }

                    dbks.put("DocNo" + (cnt > 9 ? cnt : "0" + cnt), (mdno != null  ? mdno : ""));
                    dbks.put("RevNo" + (cnt > 9 ? cnt : "0" + cnt), (mdrn != null  ? mdrn : ""));
                    dbks.put("Title" + (cnt > 9 ? cnt : "0" + cnt), xdoc.getDisplayName());
                    dbks.put("Task" + (cnt > 9 ? cnt : "0" + cnt), task.getName());
                    dbks.put("DocName" + (cnt > 9 ? cnt : "0" + cnt), (mdnm != null  ? mdnm : ""));
                    dbks.put("ReceivedOn" + (cnt > 9 ? cnt : "0" + cnt), (rcvo != null ? rcvo : ""));
                    dbks.put("ProcessTitle" + (cnt > 9 ? cnt : "0" + cnt), (processInstance != null ? processInstance.getDisplayName() : ""));
                    dbks.put("ProjectNo" + (cnt > 9 ? cnt : "0" + cnt), (prjn != null  ? prjn : ""));
                    dbks.put("DoxisDocLink" + (cnt > 9 ? cnt : "0" + cnt), mcfg.getString("webBase") + helper.getTaskURL(task.getID()));
                    cnt++;

                }

                dbks.put("docs", String.join(", ", docs));

                String tplMailPath = Utils.exportDocument(mtpl, Conf.DeleteProcess.MainPath, mtpn + "[" + uniqueId + "]");
                String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, 0,
                        Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
                );
                String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                        Conf.DeleteProcessSheetIndex.Deletion,
                        Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

                JSONObject mail = new JSONObject();
                mail.put("To", String.join(";", mails));
                //mail.put("Subject", "New Documents Deletion Request");
                mail.put("Subject", "Doc.(s) Deletion Request");
                mail.put("BodyHTMLFile", mailHtmlPath);
                Utils.sendHTMLMail(mail, mcfg);
            }
        } catch (Exception e) {
            log.error("Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        log.info("----DeleteDocument Agent Finished -----:" + task.getID());
        return resultSuccess("Agent Finished Succesfully");
    }

}
