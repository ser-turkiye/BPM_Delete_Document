package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.Objects;

public class DeleteDocument extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    private IDocument mainDocument;
    private ITask mainTask;
    private String notes;
    String prjCode = "";
    @Override
    protected Object execute() {
        mainTask = getEventTask();
        if (mainTask == null) return resultError("OBJECT CLIENT ID is NULL or not of Type ITask");
        try {
            log.info("----DeleteDocument Agent Started -----");
            this.helper = new ProcessHelper(getSes());
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
                this.deleteDocument(xdoc);
            }
            mainTask.commit();
        } catch (Exception e) {
            log.error("Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        return resultSuccess("Agent Finished Succesfully");
    }

    private void deleteDocument(IInformationObject mainInfo) throws Exception {
        mainDocument = (IDocument) mainInfo;
        String mainDocInfo = mainDocument.getDisplayName();
        log.info("Deleting Document :" + mainDocument.getID());
        try {
            ILink[] links = getDocumentServer().getReferencedRelationships(getSes(),mainDocument,false,false);
            for (ILink link : links) {
                IInformationObject xdoc = link.getTargetInformationObject();
                String docInfo = xdoc.getDisplayName();
                //if (!xdoc.getClassID().equals(Conf.ClassIDs.EngineeringDocument)){continue;}
                //getDocumentServer().deleteDocument(getSes(),mainDocument);
                getDocumentServer().deleteInformationObject(getSes(),xdoc);
                mainTask.setDescriptorValue("Notes",(Objects.equals(notes, "") ? "Deleted InformationObject :" + docInfo : notes + "\n" + "Deleted InformationObject :" + docInfo));
            }
            getDocumentServer().deleteDocument(getSes(),mainDocument);
            mainTask.setDescriptorValue("Notes",(Objects.equals(notes, "") ? "Deleted InformationObject :" + mainDocInfo : notes + "\n" + "Deleted InformationObject :" + mainDocInfo));
            log.info("Deleted Document");
        } catch (Exception e) {
            throw new Exception("Exeption Caught..deleteDocument: " + e);
        }
    }

}
