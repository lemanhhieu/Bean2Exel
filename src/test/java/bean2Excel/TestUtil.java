package bean2Excel;

public class TestUtil {
    public static long measureSpeed(String testName, Func func) throws Exception{
        System.out.println("Testing speed for \"" + testName + "\"");

        long start = System.currentTimeMillis();
        func.exec();
        long end = System.currentTimeMillis();
        long duration = end - start;
        System.out.println("Operation duration: " + duration);
        return duration;
    }

    public interface Func {
        void exec() throws Exception;
    }

}
