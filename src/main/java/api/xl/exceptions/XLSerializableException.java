package api.xl.exceptions;

public final class XLSerializableException extends Exception{
    public XLSerializableException() {
        super("Object is not XLSerializable, process is not allowed.");
    }
    public XLSerializableException(String message) {
        super(message);
    }
}
