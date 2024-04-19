package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class DeletionDocs extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    IProcessInstance processInstance;
    IInformationObject projectInfObj;
    JSONObject projects = new JSONObject();
    private ProcessHelper helper;
    private IDocument mainDocument;
    private ITask task;
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

                String tplMailPath = Utils.exportDocument(mtpl, Conf.DeleteProcess.MainPath, mtpn + "[" + uniqueId + "]");
                String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, Conf.DeleteProcessSheetIndex.Deletion,
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
