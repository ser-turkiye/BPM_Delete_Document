package ser;

import ser.ProcessHelper;
import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import com.ser.blueline.metaDataComponents.IArchiveClass;
import com.ser.blueline.metaDataComponents.IArchiveFolderClass;
import com.ser.foldermanager.IFolder;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.Arrays;
import java.util.List;

public class ProcessHelper {

    private IDocumentServer documentServer;
    private ISession session;

    private Logger log = LogManager.getLogger();
    public ProcessHelper(ISession session){
        this.session = session;
        this.documentServer = session.getDocumentServer();
    }

    public IProcessInstance buildNewProcessInstanceForID(String id){
        try{
            log.info("Building new Process Instance with ID: " +id );
            if(id ==null) return null;
            IBpmService bpm = session.getBpmService();
            IProcessType processType = bpm.getProcessType(id);
            if(processType ==null) {
                log.error("Process Type with ID couldn't be found");
                return null;
            }
            IBpmService.IProcessInstanceBuilder processBuilder = bpm.buildProcessInstanceObject().setProcessType(processType);
            IBpmDatabase processDB = session.getBpmDatabaseByName(processType.getDefaultDatabase().getName());
            if(processDB ==null){
                log.error("Process Type: " + processType.getName() + " has no assigned databases");
                return null;
            }
            processBuilder.setBpmDatabase(processDB);
            processBuilder.setValidationProcessDefinition(processType.getActiveProcessDefinition());
            return processBuilder.build();
        }catch(Exception e){
            log.error(e.getMessage());
            return null;
        }
    }
    public boolean mapDescriptorsFromObjectToObject(IInformationObject srcObject , IInformationObject targetObject , boolean overwriteValues){
        log.info("Mapping Descriptors from PInformation to Information Object");
        String[] targeObjectAssignedDesc;
        if(targetObject instanceof IFolder){
            log.info("Information Object is of type IFolder");
            String classID = targetObject.getClassID();
            IArchiveFolderClass folderClass = documentServer.getArchiveFolderClass(classID , session);
            targeObjectAssignedDesc = folderClass.getAssignedDescriptorIDs();
        }else if(targetObject instanceof IDocument){
            log.info("Information Object is of type IDocument");
            IArchiveClass documentClass = ((IDocument)targetObject).getArchiveClass();
            targeObjectAssignedDesc = documentClass.getAssignedDescriptorIDs();
        }else if(targetObject instanceof ITask){
            log.info("Information Object is of type ITask");
            IProcessType processType = ((ITask)targetObject).getProcessType();
            targeObjectAssignedDesc = processType.getAssignedDescriptorIDs();
        }else if(targetObject instanceof IProcessInstance){
            log.info("Information Object is of type IProcessInstace");
            IProcessType processType = ((IProcessInstance)targetObject).getProcessType();
            targeObjectAssignedDesc = processType.getAssignedDescriptorIDs();
        }else{
            log.error("Information Object is not of Supported type");
            return false;
        }
        List<String> targetDesc = Arrays.asList(targeObjectAssignedDesc);
        IValueDescriptor[] srcDesc = srcObject.getDescriptorList();

        for(int i=0; i <  srcDesc.length; i++){
            IValueDescriptor vd = srcDesc[i];
            String descID = vd.getId();
            String descName = vd.getName();
            int descType = ((IDescriptor) vd).getMultiValueType();
            try{
                if(targetDesc.contains(descID)){
                    if(targetObject.getDescriptorValue(descID) != null && targetObject.getDescriptorValue(descID) != "")
                        if(!overwriteValues)continue;

                    if(((IDescriptor) vd).getMultiValueType() != ((IDescriptor) vd).MULTI_VALUE_TYPE_NONE){
                        targetObject.setDescriptorValues(descID,srcObject.getDescriptorValues(descID,String.class));
                    }else {
                        targetObject.setDescriptorValue(descID, srcObject.getDescriptorValue(descID));
                    }
                }
            } catch (Exception e) {
                log.error("Exception caught while adding descriptor: "+descName);
                log.error(e.getMessage());
                return false;
            }
        }
        return true;
    }
    public String getDocumentURL(String documentID){
        StringBuilder webcubeUrl = new StringBuilder();
        webcubeUrl.append("?system=").append(session.getSystem().getName());
        webcubeUrl.append("&action=showdocument&home=1&reusesession=1&id=").append(documentID);
        return webcubeUrl.toString();
    }
    public String getTaskURL(String taskID){
        StringBuilder webcubeUrl = new StringBuilder();
        webcubeUrl.append("?system=").append(session.getSystem().getName());
        webcubeUrl.append("&action=showtask&home=1&reusesession=1&id=").append(taskID);
        return webcubeUrl.toString();
    }
}
