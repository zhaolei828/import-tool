package sample.freemarker;

public class WriteIntoExcelException extends LeiRuntimeException
{

    private static final long serialVersionUID = 1336373165807040572L;

    public WriteIntoExcelException(String errorCode, String message)
    {
        super(errorCode, message);
    }

    public WriteIntoExcelException(String errorCode, String message, String desc) {
        super(errorCode, message, desc);
    }

    public WriteIntoExcelException(String errorCode, Throwable ex)
    {
        super(errorCode, ex);
    }

    public WriteIntoExcelException(String errorCode, String message, Throwable e)
    {
        super(errorCode, message, e);
    }

    public String toString()
    {
        return super.toString();
    }

}
