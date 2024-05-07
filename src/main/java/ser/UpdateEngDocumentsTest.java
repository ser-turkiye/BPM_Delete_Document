package ser;

import com.ser.blueline.IDocument;
import com.ser.blueline.IInformationObject;
import com.ser.blueline.ILink;
import com.ser.blueline.InformationObjectType;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.TaskStatus;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.TimeUnit;

public class UpdateEngDocumentsTest extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    List<String> updated = new ArrayList<>();
    String decisionCode = "";
    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        try {
            helper = new ProcessHelper(getSes());
            log.info("-- UpdateEngDocuments Agent Started -----");
            IInformationObject[] list = this.getEmptyDocIDEngineeringDocs(helper);
            this.updateDocIDEngDocs(list);
        } catch (Exception e) {
            log.error("UpdateEngDocuments Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        log.info("-- UpdateEngDocuments Agent Finished -----");
        return resultSuccess("Agent Finished Succesfully");
    }
    public void updateDecisionCodeEngDocs(IInformationObject[] list) throws Exception {
        for(IInformationObject infoDoc : list){
            ILink[] links2 = getDocumentServer().getReferencingRelationships(getSes(),((IDocument) infoDoc).getID(),false);
            for (ILink link2 : links2) {
                IInformationObject xdoc = link2.getSourceInformationObject();
                InformationObjectType objType = xdoc.getInformationObjectType();
                log.info("usage object : " + xdoc.getID() + " /// type : " + objType);
                if(objType != InformationObjectType.PROCESS_INSTANCE){continue;}
                if(Objects.equals(xdoc.getClassID(), Conf.ClassIDs.ReviewMain)){
                    decisionCode = "";
                    IProcessInstance proi = (IProcessInstance) xdoc;
                    IDocument mainDocument = (IDocument) proi.getMainInformationObject();
                    if(mainDocument == null){
                        log.info("No-Main document");
                    }
                    Collection<ITask> tsks = proi.findTasks();
                    for(ITask ttsk : tsks) {
                        if (ttsk.getStatus() != TaskStatus.COMPLETED) {
                            continue;
                        }
                        if(decisionCode != ""){break;}
                        String tnam = (ttsk.getName() != null ? ttsk.getName() : "");
                        String tcod = (ttsk.getCode() != null ? ttsk.getCode() : "");
                        log.info("TASK-Name[" + tnam + "]");
                        log.info("TASK-Code[" + tcod + "]");
                        if(ttsk.getLoadedParentTask() != null && (tnam.equals("Consolidator Review") || tcod.equals("Step03"))){
                            decisionCode = ttsk.getDecision().getCode();
                            log.info("Approval Code updated.. decisionCode IS:" + decisionCode);
                            mainDocument.setDescriptorValue("ccmPrjDocApprCode", decisionCode);
                            mainDocument.commit();
                            updated.add(mainDocument.getID() + " > " + mainDocument.getDescriptorValue(Conf.Descriptors.DocNumber) + " > " + decisionCode);
                            log.info("UpdateEngDocument batch...maindoc committed...APPR CODE IS:" + mainDocument.getDescriptorValue("ccmPrjDocApprCode"));
                            TimeUnit.SECONDS.sleep(30);
                            log.info("UpdateEngDocument batch...sleeping (30 second)");
                        }
                    }
                }
            }
        }
    }
    public IInformationObject[] getEngineeringDocs(ProcessHelper helper) {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append("CCMPRJDOCWFTASKNAME").append(" = '").append("Completed").append("'")
                .append(" AND ")
                .append("CCMPRJDOCAPPRCODE").append(" IS NULL");
        String whereClause = builder.toString();
        log.info("Where Clause: " + whereClause);
        IInformationObject[] list = helper.createQuery(new String[]{Conf.Databases.EngineeringDocument}, whereClause, "", 0, false);
        return list;
    }
    public void updateDocIDEngDocs(IInformationObject[] list) throws Exception {
        for(IInformationObject infoDoc : list){
            IDocument mainDocument = (IDocument) infoDoc;
            String docID = mainDocument.getID();
            log.info("DOC ID updated.. doc ID:" + docID);
            mainDocument.setDescriptorValue("DocID", docID);
            mainDocument.commit();
            log.info("Updated EngDocument batch...maindoc committed...DOC ID IS:" + mainDocument.getDescriptorValue("DocID"));
            TimeUnit.SECONDS.sleep(30);
            log.info("UpdateEngDocument batch...sleeping (30 second)");
        }
    }
    public IInformationObject[] getEmptyDocIDEngineeringDocs(ProcessHelper helper) {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.EngineeringDocument).append("'")
                .append(" AND ")
                .append("DOCID").append(" IS NULL");
        String whereClause = builder.toString();
        log.info("Where Clause: " + whereClause);
        IInformationObject[] list = helper.createQuery(new String[]{Conf.Databases.EngineeringDocument}, whereClause, "", 0, false);
        return list;
    }
}
