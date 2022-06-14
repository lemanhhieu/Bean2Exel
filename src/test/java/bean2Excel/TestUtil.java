package bean2Excel;

public class TestUtil {
    public static long measureSpeed(Func func) throws Exception{
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
