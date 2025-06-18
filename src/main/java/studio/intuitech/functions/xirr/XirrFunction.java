package studio.intuitech.functions.xirr;

import org.apache.poi.ss.formula.LazyRefEval;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.functions.MultiOperandNumericFunction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import studio.intuitech.event.Event;
import studio.intuitech.event.Handler;
import studio.intuitech.functions.newtonsmethod.NewtonConfig;
import studio.intuitech.functions.newtonsmethod.NewtonsMethod;

public class XirrFunction implements FreeRefFunction, Handler<Event> {

    private final Workbook workbook;
    private boolean enable = true;

    public XirrFunction(Workbook workbook) {
        this.workbook = workbook;
    }

    static final void checkValue(double result) throws EvaluationException {
        if (Double.isNaN(result) || Double.isInfinite(result)) {
            throw new EvaluationException(ErrorEval.NUM_ERROR);
        }
    }

    public void setEnable(boolean enable) {
        this.enable = enable;
    }

    @Override
    public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
        if (!enable) {
            return ErrorEval.NA;
        }
        if (args.length < 2 || args.length > 4) {
            return ErrorEval.VALUE_INVALID;
        }
        Cell targetCell = null;
        double result;

        try {

            double[] values = ValueCollector.collectValues(args[0]);
            double[] dates = ValueCollector.collectValues(args[1]);

            double guess;
            if (args.length == 3) {
                ValueEval v = OperandResolver.getSingleValue(args[2], ec.getRowIndex(), ec.getColumnIndex());
                guess = OperandResolver.coerceValueToDouble(v);
            } else {
                guess = 0.1d;
            }

            Sheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());


            if (args.length == 4) {
                ValueEval targetCellVE = args[3];
                targetCell = sheet.getRow(((LazyRefEval) targetCellVE).getRow()).getCell(((LazyRefEval) targetCellVE).getColumn());
            }
            result = calculateXIRR(values, dates, guess);

            checkValue(result);

        } catch (EvaluationException e) {
            //e.printStackTrace();
            return e.getErrorEval();
        }
        System.out.println(result);
        if (targetCell != null) {
            targetCell.setCellValue(result);
        }
        return new NumberEval(result);
    }

    public Double calculateXIRR(double[] values, double[] dates, double guess) {

        double tolerance = 1e-3;
        double epsilon = Math.sqrt(Math.ulp(1d));
        int maxIterations = 1000;

        NewtonConfig config = new NewtonConfig(tolerance, epsilon, maxIterations);
        XirrContext context = new XirrContext(values, dates, guess);

        return NewtonsMethod.newtonsMethod(0d, 0d, context, Xirr.total_f_xirr, Xirr.total_df_xirr, config);

    }

    @Override
    public void onEvent(Event event) {
        if (event.equals(Event.GOAL_SEEK_DONE)) {
            setEnable(true);
        }
    }

    static final class ValueCollector extends MultiOperandNumericFunction {
        private static final ValueCollector instance = new ValueCollector();

        public ValueCollector() {
            super(false, false);
        }

        public static double[] collectValues(ValueEval... operands) throws EvaluationException {
            return instance.getNumberArray(operands);
        }

        protected double evaluate(double[] values) {
            throw new IllegalStateException("should not be called");
        }
    }
}
