package studio.intuitech.functions.goalseek;

import java.time.LocalDateTime;
import java.util.function.BiFunction;

import org.apache.poi.ss.formula.LazyRefEval;
import org.apache.poi.ss.formula.OperationEvaluationContext;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import studio.intuitech.event.Event;
import studio.intuitech.event.EventDispatcher;
import studio.intuitech.functions.newtonsmethod.NewtonConfig;
import studio.intuitech.functions.newtonsmethod.NewtonsMethod;

public class GoalSeekFunction implements FreeRefFunction {

    private final Workbook workbook;

    private final EventDispatcher eventDispatcher;

    private boolean firstUse = true;

    public GoalSeekFunction(Workbook workbook, EventDispatcher eventDispatcher) {
        this.workbook = workbook;
        this.eventDispatcher = eventDispatcher;
    }

    private static Double calculateGoalSeek(FormulaEvaluator formulaEvaluator, Cell changingCell, Cell targetCell, double estimationValue, double targetCellValue) {
        BiFunction<GoalSeekContext, Double, Double> f = (context, x) -> {
            formulaEvaluator.clearAllCachedResultValues();
            changingCell.setCellValue(x);
            formulaEvaluator.evaluateAll();

            return targetCell.getNumericCellValue();
        };

        double tolerance = 1e-3;
        double epsilon = Math.sqrt(Math.ulp(1d));
        int maxIterations = 1000;
        NewtonConfig config = new NewtonConfig(tolerance, epsilon, maxIterations);

        BiFunction<GoalSeekContext, Double, Double> df = (context, y) -> (f.apply(context, y + epsilon) - f.apply(context, y)) / epsilon;
        // (f2 - f1) / epsilon; Deriválás

        System.out.println(LocalDateTime.now());

        return NewtonsMethod.newtonsMethod(
            estimationValue,
            targetCellValue,
            new GoalSeekContext(),
            f,
            df,
            config);
    }

    @Override
    public ValueEval evaluate(ValueEval[] args, OperationEvaluationContext ec) {
        if (args.length != 4) {
            return ErrorEval.VALUE_INVALID;
        }

        if (firstUse) {
            firstUse = false;
        } else {
            return ErrorEval.CIRCULAR_REF_ERROR;
        }

        ValueEval targetCellVE = args[0];
        ValueEval targetValueCellVE = args[1];
        ValueEval changingCellVE = args[2];
        ValueEval estimationCellVE = args[3];

        try {
            double targetCellValue = OperandResolver.coerceValueToDouble(
                OperandResolver.getSingleValue(targetValueCellVE, ec.getRowIndex(), ec.getColumnIndex()));
            double estimationValue = OperandResolver.coerceValueToDouble(
                OperandResolver.getSingleValue(estimationCellVE, ec.getRowIndex(), ec.getColumnIndex()));


            Sheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

            final Cell changingCell = sheet.getRow(((LazyRefEval) changingCellVE).getRow()).getCell(((LazyRefEval) changingCellVE).getColumn());
            final Cell targetCell = sheet.getRow(((LazyRefEval) targetCellVE).getRow()).getCell(((LazyRefEval) targetCellVE).getColumn());

            FormulaEvaluator formulaEvaluator = this.workbook.getCreationHelper().createFormulaEvaluator();

            Double result = calculateGoalSeek(formulaEvaluator, changingCell, targetCell, estimationValue, targetCellValue);

            if (result != null) {
                System.out.println("A megoldás: " + result);
                eventDispatcher.dispatch(Event.GOAL_SEEK_DONE);

                formulaEvaluator.clearAllCachedResultValues();
                changingCell.setCellValue(result);

                formulaEvaluator.evaluateAll();
            } else {
                System.out.println("Az iterációk során nem sikerült megoldást találni.");
            }
            eventDispatcher.dispatch(Event.GOAL_SEEK_DONE);

            System.out.println(LocalDateTime.now());


            return new NumberEval(targetCellValue);
        } catch (EvaluationException e) {
            e.printStackTrace();
            return e.getErrorEval();
        }
    }

}
