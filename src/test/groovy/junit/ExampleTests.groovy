package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.DeleteDocument

class ExampleTests {

    Binding binding

    @BeforeClass
    static void initSessionPool() {
        AgentTester.initSessionPool()
    }

    @Before
    void retrieveBinding() {
        binding = AgentTester.retrieveBinding()
    }

    @Test
    void testForAgentResult() {

        def agent = new DeleteDocument();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM24857a0b4a-96eb-4bc4-9449-7bc5f7fb7d9d182024-02-06T07:29:10.382Z011"

        def result = (AgentExecutionResult)agent.execute(binding.variables)
        System.out.println(result)
//        assert result.resultCode == 0
//        assert result.executionMessage.contains("Linux")
//        assert agent.eventInfObj instanceof IDocument
    }

    @Test
    void testForGroovyAgentMethod() {
//        def agent = new GroovyAgent()
//        agent.initializeGroovyBlueline(binding.variables)
//        assert agent.getServerVersion().contains("Linux")
    }

    @Test
    void testForJavaAgentMethod() {
//        def agent = new JavaAgent()
//        agent.initializeGroovyBlueline(binding.variables)
//        assert agent.getServerVersion().contains("Linux")
    }

    @After
    void releaseBinding() {
        println("RLEASE BINDING RUNNING.....")
        AgentTester.releaseBinding(binding)
    }

    @AfterClass
    static void closeSessionPool() {
        AgentTester.closeSessionPool()
    }
}
