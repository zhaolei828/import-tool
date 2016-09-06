package sample.freemarker;

public class BusinessException extends Exception
{
    private static final long serialVersionUID = 1498884168645406117L;

    public BusinessException(String message)
    {
        super(message);
    }
}
