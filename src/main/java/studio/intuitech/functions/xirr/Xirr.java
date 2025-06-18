package studio.intuitech.functions.xirr;

import java.util.function.BiFunction;

/*
 Based on:
 Class Xirr from https://stackoverflow.com/questions/36789967/java-program-to-calculate-xirr-without-using-excel-or-any-other-library
*/
final class Xirr {

    public static BiFunction<XirrContext, Double, Double> total_f_xirr = (XirrContext context, Double x0) -> {
        double resf = 0d;
        for (int i = 0; i < context.payments().length; i++) {
            resf = resf + f_xirr(context.payments()[i], context.days()[i], context.days()[0], x0);
        }

        return resf;
    };
    public static BiFunction<XirrContext, Double, Double> total_df_xirr = (XirrContext context, Double x0) -> {
        double resf = 0d;
        for (int i = 0; i < context.payments().length; i++) {
            resf = resf + df_xirr(context.payments()[i], context.days()[i], context.days()[0], x0);
        }
        return resf;
    };

    private static double f_xirr(double p, double dt, double dt0, double x) {
        return p * Math.pow((x + 1d), (dt0 - dt) / 365d);
    }

    private static double df_xirr(double p, double dt, double dt0, double x) {
        return (1d / 365d) * (dt0 - dt) * p * Math.pow((x + 1d), ((dt0 - dt) / 365d) - 1d);
    }
}