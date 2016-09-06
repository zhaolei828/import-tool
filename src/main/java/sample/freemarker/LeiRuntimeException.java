package sample.freemarker;

public class LeiRuntimeException extends RuntimeException
{
    private static final long serialVersionUID = -7976146395762011863L;

    private String message;

    private String errorCode = null;

    private String desc;

    public String getErrorCode()
    {
        return errorCode;
    }

    public void setErrorCode(String errorCode)
    {
        this.errorCode = errorCode;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public LeiRuntimeException()
    {
        super();
    }

    public LeiRuntimeException(String errorCode, String message)
    {
        super(message);
        this.errorCode = errorCode;
        this.message = message;
    }

    public LeiRuntimeException(String errorCode, String message,String desc)
    {
        super(message);
        this.desc = desc;
        this.errorCode = errorCode;
        this.message = message;
    }

    public LeiRuntimeException(String errorCode, Throwable ex)
    {
        super(ex);
        this.errorCode = errorCode;
    }

    public LeiRuntimeException(String errorCode, String message, Throwable e)
    {
        super(message, e);
        this.errorCode = errorCode;
        this.message = message;
    }

    public String getMessage()
    {
        return this.message;
    }

    public String toString()
    {
        StringBuffer buf = new StringBuffer();
        buf.append(super.toString());
        if(null != errorCode)
        {
            buf.append("<" + errorCode + ">");
        }
        return buf.toString();
    }
}
