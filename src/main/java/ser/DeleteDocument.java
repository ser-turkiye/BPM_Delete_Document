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
import java.util.Objects;
import java.util.UUID;

public class DeleteDocument extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    private IDocument mainDocument;
    private ITask mainTask;
    private String notes;
    String prjn = "";
    List<String> docs = new ArrayList<>();
    @Override
    protected Object execute() {
        mainTask = getEventTask().findParentTask();
        IProcessInstance processInstance = mainTask.getProcessInstance();
        String currentUser = getSes().getUser().getFullName();
        IUser processOwner = processInstance.getOwner();
        String uniqueId = UUID.randomUUID().toString();

        if (mainTask == null) return resultError("OBJECT CLIENT ID is NULL or not of Type ITask");
        try {
            com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);
            this.helper = new ProcessHelper(getSes());
            log.info("----DeleteDocumentProcess Agent Started -----:" + mainTask.getID());
            notes = mainTask.getDescriptorValue("Notes");
            if(mainTask.getProcessInstance().findLockInfo().getOwnerID() != null){
                log.error("Task is locked.." + mainTask.getID() + "..restarting agent");
                return resultRestart("Restarting Agent");
            }

            IInformationObjectLinks links = mainTask.getProcessInstance().getLoadedInformationObjectLinks();
            for (ILink link : links.getLinks()) {
                IInformationObject xdoc = link.getTargetInformationObject();
                String taskName = xdoc.getDescriptorValue("ccmPrjDocWFTaskName");
                if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                prjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                this.deleteDocument(xdoc);
            }
            //mainTask.commit();
            String mtpn = "DOCUMENT_DELETION_MAIL";
            JSONObject dbks = new JSONObject();
            dbks.put("docs", String.join(", ", docs));

            IDocument mtpl = Utils.getTemplateDocument(prjn, mtpn, helper);
            if(mtpl == null){
                log.info("Template-Document [ " + mtpn + " ] not found.");
                //throw new Exception("Template-Document [ " + mtpn + " ] not found.");
            }else {
                String tplMailPath = Utils.exportDocument(mtpl, Conf.DeleteProcess.MainPath, mtpn + "[" + uniqueId + "]");
                String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, Conf.DeleteProcessSheetIndex.Mail,
                        Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
                );
                String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

                String umail = processOwner.getEMailAddress();
                List<String> mails = new ArrayList<>();
                log.info("Mail To : " + String.join(";", mails));
                if (umail != null) {
                    mails.add(umail);
                    JSONObject mail = new JSONObject();
                    mail.put("To", String.join(";", mails));
                    mail.put("Subject", "Deletion Documents");
                    mail.put("BodyHTMLFile", mailHtmlPath);
                    Utils.sendHTMLMail(getSes(), getSes().getDocumentServer(), "CCM_MAIL_CONFIG", mail);
                } else {
                    log.info("Mail adress is null :" + processOwner.getFullName());
                }
            }
        } catch (Exception e) {
            log.error("Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        log.info("----DeleteDocument Agent Finished -----:" + mainTask.getID());
        return resultSuccess("Agent Finished Succesfully");
    }
    private void deleteDocument(IInformationObject mainInfo) throws Exception {
        mainDocument = (IDocument) mainInfo;
        String mainDocInfo = mainDocument.getDisplayName();
        log.info("Deleting Document :" + mainDocument.getID());
        docs.add(mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber));
        try {
            ILink[] links = getDocumentServer().getReferencedRelationships(getSes(),mainDocument,false,false);
            ILink[] links2 = getDocumentServer().getReferencingRelationships(getSes(),mainDocument.getID(),false);
            for (ILink link : links) {
                IInformationObject xdoc = link.getTargetInformationObject();
                String docInfo = xdoc.getDisplayName();
                log.info("Delete link doc : " + xdoc.getID());
                getDocumentServer().deleteInformationObject(getSes(),xdoc);
                log.info("deleted link doc");
                //mainTask.setDescriptorValue("Notes",(Objects.equals(notes, "") ? "Deleted InformationObject :" + docInfo : notes + "\n" + "Deleted InformationObject :" + docInfo));
            }
            for (ILink link2 : links2) {
                IInformationObject xdoc = link2.getSourceInformationObject();
                String docClassID = xdoc.getClassID();
                InformationObjectType objType = xdoc.getInformationObjectType();
                log.info("Delete usage object : " + xdoc.getID() + " /// type : " + objType);
                if(objType != InformationObjectType.PROCESS_INSTANCE){continue;}
                //if(Objects.equals(docClassID, Conf.ClassIDs.ProjectCard)){continue;}
                getDocumentServer().deleteInformationObject(getSes(),xdoc);
                log.info("deleted usage object");
                //mainTask.setDescriptorValue("Notes",(Objects.equals(notes, "") ? "Deleted InformationObject :" + docInfo : notes + "\n" + "Deleted InformationObject :" + docInfo));
            }
            getDocumentServer().deleteDocument(getSes(),mainDocument);
            //mainTask.setDescriptorValue("Notes",(Objects.equals(notes, "") ? "Deleted InformationObject :" + mainDocInfo : notes + "\n" + "Deleted InformationObject :" + mainDocInfo));
            log.info("Deleted Main Document");
        } catch (Exception e) {
            throw new Exception("Exeption Caught..deleteDocument: " + e);
        }
    }

}
