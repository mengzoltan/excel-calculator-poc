package studio.intuitech.event;

import java.util.HashMap;
import java.util.Map;

public class EventDispatcher {

    private final Map<Class<? extends Event>, Handler<? extends Event>> handlers;

    public EventDispatcher() {
        handlers = new HashMap<>();
    }

    /**
     * Links an {@link Event} to a specific {@link Handler}.
     *
     * @param eventType The {@link Event} to be registered
     * @param handler   The {@link Handler} that will be handling the {@link Event}
     */
    public <E extends Event> void registerHandler(Class<E> eventType, Handler<E> handler) {
        handlers.put(eventType, handler);
    }

    /**
     * Dispatches an {@link Event} depending on its type.
     *
     * @param event The {@link Event} to be dispatched
     */
    @SuppressWarnings("unchecked")
    public <E extends Event> void dispatch(E event) {
        var handler = (Handler<E>) handlers.get(event.getClass());
        if (handler != null) {
            handler.onEvent(event);
        }
    }
}
