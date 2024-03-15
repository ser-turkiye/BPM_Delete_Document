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

public class DeleteMultiReviewProcess extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    private ProcessHelper helper;
    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();

        try {
            helper = new ProcessHelper(getSes());
            log.info("-- Delete MultiReview Process Agent Started -----");
            //projects = Utils.getProjectWorkspaces(helper);
            IInformationObject[] multiprcss = Utils.getMultiReviewProcesses(helper);
            for(IInformationObject info : multiprcss){
                IProcessInstance pi = ((ITask) info).getProcessInstance();
                log.info("--- Delete MultiReview Process ID -----" + pi.getID());
                if(pi.delete()){
                    log.info("---- Deleted -----");
                }else{
                    log.info("---- Not Deleted -----");
                }
            }
        } catch (Exception e) {
            log.error("Delete MultiReview Exception Caught");
            log.error(e.getMessage());
            return resultError(e.getMessage());
        }
        log.info("-- Delete MultiReview Agent Finished -----");
        return resultSuccess("Agent Finished Succesfully");
    }
}
