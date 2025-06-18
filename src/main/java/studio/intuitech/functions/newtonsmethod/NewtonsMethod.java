package studio.intuitech.functions.newtonsmethod;

import java.util.function.BiFunction;

public class NewtonsMethod {

    public static <T> Double newtonsMethod(
        double x0,
        double target,
        T context,
        BiFunction<T, Double, Double> f,
        BiFunction<T, Double, Double> df,
        NewtonConfig config
    ) {

        double x1;
        for (int i = 0; i < config.maxIterations(); i++) {
            double y = f.apply(context, x0) - target;
            double dy = df.apply(context, x0);

            if (Math.abs(dy) < config.epsilon()) {  // Avoid division by a very small number
                return null;
            }

            x1 = x0 - y / dy;  // Newton's method calculation

            if (Math.abs(x1 - x0) <= config.tolerance()) {  // Converged within tolerance
                return x1;
            }

            x0 = x1;  // Update x0 for next iteration
        }
        return null;  // Did not converge
    }
}