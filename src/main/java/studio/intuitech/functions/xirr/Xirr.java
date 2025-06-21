package studio.intuitech.functions.xirr;

import java.util.function.BiFunction;

final class Xirr {

    public static BiFunction<XirrContext, Double, Double> f = (context, rate) -> {
        double sum = 0.0;
        double[] values = context.payments();
        double[] dates = context.days();
        double date0 = dates[0];

        for (int i = 0; i < values.length; i++) {
            double dt = (dates[i] - date0) / 365.0;
            sum += values[i] / Math.pow(1.0 + rate, dt);
        }
        return sum;
    };

    public static BiFunction<XirrContext, Double, Double> df = (context, rate) -> {
        double sum = 0.0;
        double[] values = context.payments();
        double[] dates = context.days();
        double date0 = dates[0];

        for (int i = 0; i < values.length; i++) {
            double dt = (dates[i] - date0) / 365.0;
            sum -= dt * values[i] / Math.pow(1.0 + rate, dt + 1);
        }
        return sum;
    };

}