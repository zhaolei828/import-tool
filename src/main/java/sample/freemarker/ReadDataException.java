package sample.freemarker;

public class ReadDataException extends LeiRuntimeException
{

    private static final long serialVersionUID = 1336373165807040572L;

    public ReadDataException(String errorCode, String message)
    {
        super(errorCode, message);
    }

    public ReadDataException(String errorCode, String message, String desc) {
        super(errorCode, message, desc);
    }

    public ReadDataException(String errorCode, Throwable ex)
    {
        super(errorCode, ex);
    }

    public ReadDataException(String errorCode, String message, Throwable e)
    {
        super(errorCode, message, e);
    }

    public String toString()
    {
        return super.toString();
    }

}
