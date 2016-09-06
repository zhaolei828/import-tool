package sample.freemarker;

public class ParseHtmlException extends LeiRuntimeException
{

    private static final long serialVersionUID = 1436373165807040572L;

    public ParseHtmlException(String errorCode, String message)
    {
        super(errorCode, message);
    }

    public ParseHtmlException(String errorCode, Throwable ex)
    {
        super(errorCode, ex);
    }

    public ParseHtmlException(String errorCode, String message, Throwable e)
    {
        super(errorCode, message, e);
    }

    public String toString()
    {
        return super.toString();
    }

}
