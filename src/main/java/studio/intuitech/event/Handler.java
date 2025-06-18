package studio.intuitech.event;

public interface Handler<E extends Event> {
    void onEvent(E event);
}
