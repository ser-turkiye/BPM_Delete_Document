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
    JSONObject projects = new JSONObject();
    private ITask mainTask;
    private String notes;
    String prjn = "";
    List<String> docs = new ArrayList<>();
    @Override
    protected Object execute() {
        helper = new ProcessHelper(getSes());
        if(getEventDocument() != null && Objects.equals(getEventDocument().getClassID(), Conf.ClassIDs.EngineeringCopy)){
            mainDocument = getEventDocument();
            try {
                ///deleting eng. copy and linked sub review processes
                log.info("----Delete EngCopyDocument Agent Started -----:" + mainDocument.getID());
                this.deleteDocument(mainDocument);
                log.info("----Delete EngCopyDocument Agent Finished -----");
            } catch (Exception e) {
                log.error("Exception Caught");
                log.error(e.getMessage());
                return resultError(e.getMessage());
            }
        }else {
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
                Utils.loadDirectory(Conf.Paths.MainPath);

                //helper = new ProcessHelper(getSes());
                log.info("----DeleteDocument Process Agent Started -----:" + mainTask.getID());
                notes = mainTask.getDescriptorValue("Notes");
                if (mainTask.getProcessInstance().findLockInfo().getOwnerID() != null) {
                    log.error("Task is locked.." + mainTask.getID() + "..restarting agent");
                    return resultRestart("Restarting Agent");
                }

                IInformationObjectLinks links = mainTask.getProcessInstance().getLoadedInformationObjectLinks();
                for (ILink link : links.getLinks()) {
                    IInformationObject xdoc = link.getTargetInformationObject();
                    String taskName = xdoc.getDescriptorValue("ccmPrjDocWFTaskName");
                    if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)) {
                        continue;
                    }
                    prjn = xdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    //this.deleteDocument(xdoc);
                    this.deleteDocument((IDocument) xdoc);
                }
                //mainTask.commit();
                String mtpn = "DOCUMENT_DELETION_MAIL";

                projects = Utils.getProjectWorkspaces(helper);
                IDocument mtpl = null;
                for (String prjn : projects.keySet()) {
                    IInformationObject prjt = (IInformationObject) projects.get(prjn);
                    IDocument dtpl = Utils.getTemplateDocument(prjt, mtpn);
                    if (dtpl == null) {
                        continue;
                    }
                    mtpl = dtpl;
                    break;
                }
                JSONObject dbks = new JSONObject();
                dbks.put("docs", String.join(", ", docs));

                if (mtpl == null) {
                    log.info("Template-Document [ " + mtpn + " ] not found.");
                    //throw new Exception("Template-Document [ " + mtpn + " ] not found.");
                } else {
                    String tplMailPath = Utils.exportDocument(mtpl, Conf.DeleteProcess.MainPath, mtpn + "[" + uniqueId + "]");
                    String mailExcelPath = Utils.saveDocReviewExcel(tplMailPath, Conf.DeleteProcessSheetIndex.Deleted,
                            Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].xlsx", dbks
                    );
                    String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                            Conf.DeleteProcessSheetIndex.Deleted,
                            Conf.DeleteProcess.MainPath + "/" + mtpn + "[" + uniqueId + "].html");

                    String umail = processOwner.getEMailAddress();
                    List<String> mails = new ArrayList<>();
                    log.info("Mail To : " + String.join(";", mails));
                    if (umail != null) {
                        mails.add(umail);
                        JSONObject mail = new JSONObject();
                        mail.put("To", String.join(";", mails));
                        mail.put("Subject", "Deleted Documents");
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
        }
        return resultSuccess("Agent Finished Succesfully");
    }
    private void deleteDocument(IDocument mainDoc) throws Exception {
        mainDocument = mainDoc;
        String prjCode = mainDocument.getDescriptorValue(Conf.Descriptors.ProjectNo);
        String docCode = mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber);
        String revCode = mainDocument.getDescriptorValue(Conf.Descriptors.Revision);
        log.info("Deleting Document :" + mainDocument.getID());
        docs.add((docCode == null ? "No Document Number" : docCode));
        try {
            ILink[] links = getDocumentServer().getReferencedRelationships(getSes(),mainDocument,false,false);
            ILink[] links2 = getDocumentServer().getReferencingRelationships(getSes(),mainDocument.getID(),false);
            for (ILink link : links) {
                IInformationObject xdoc = link.getTargetInformationObject();
                String docInfo = xdoc.getDisplayName();
                log.info("Delete link doc : " + xdoc.getID());
                getDocumentServer().deleteInformationObject(getSes(),xdoc);
                log.info("deleted link doc");
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
            if(!Objects.equals(mainDocument.getClassID(), Conf.ClassIDs.EngineeringCopy)){
                IInformationObject[] infoEngCopyDocs = Utils.getEngineeringCopyDocuments(helper,prjCode,docCode,revCode);
                for(IInformationObject infoCopyDoc : infoEngCopyDocs){
                    infoCopyDoc.setDescriptorValue("ObjectName","DELETED");
                    infoCopyDoc.commit();
                }
            }
            getDocumentServer().deleteDocument(getSes(),mainDocument);
            log.info("Deleted Document");
        } catch (Exception e) {
            throw new Exception("Exeption Caught..deleteDocument: " + e);
        }
    }
    private void deleteDocumentOLD(IInformationObject mainInfo) throws Exception {
        mainDocument = (IDocument) mainInfo;
        String apprCode = mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber);
        log.info("Deleting Document :" + mainDocument.getID());
        docs.add((apprCode == null ? "No Document Number" : apprCode));
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
