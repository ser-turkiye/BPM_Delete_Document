/*
@groovy.transform.BaseScript com.ser.blueline.groovy.CSBScript testAgent

// TestAgent.groovy is run by AgentTester, when not command parameters are supplie

// This is how logging should be done for CSBScript Agents in CSB 4.2+
log2.info("Hello from a CSBScript agent via apache.commons.logging")
// The old log4j1 Logger is deprecated but still works
log.info("Hello via log4j1")

return resultSuccess(srv.serverInfo.serverVersion)

*/