package studio.intuitech;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneOffset;

import org.apache.poi.ss.formula.functions.FreeRefFunction;
import org.apache.poi.ss.formula.udf.AggregatingUDFFinder;
import org.apache.poi.ss.formula.udf.DefaultUDFFinder;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import studio.intuitech.event.Event;
import studio.intuitech.event.EventDispatcher;
import studio.intuitech.functions.CalculationExcelTemplateFactory;
import studio.intuitech.functions.CalculationExcelType;
import studio.intuitech.functions.goalseek.GoalSeekFunction;
import studio.intuitech.functions.xirr.XirrFunction;

public class MainExcelCalculator {

    public static void main(String[] args) throws IOException {

        System.out.println("---------GOAL SEEK -----------");
        String outputNormal = "./normal_results.xlsx";
        test(CalculationExcelType.NORMAL, outputNormal);
        System.out.println();

        System.out.println("---------XIRR-----------");
        String outputXirr = "./xirr_result.xlsx";
        test(CalculationExcelType.ONLY_XIRR, outputXirr);
    }

    private static void test(CalculationExcelType excelType, String outFilePath) throws IOException {

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(CalculationExcelTemplateFactory.getCalculationExcelTemplate(excelType)))) {


            EventDispatcher eventDispatcher = new EventDispatcher();

            String[] functionNames = { "goalSeek", "myXIRR" };
            XirrFunction xirrFunction = new XirrFunction(workbook);

            eventDispatcher.registerHandler(Event.class, xirrFunction);
            if (!excelType.equals(CalculationExcelType.ONLY_XIRR)) {
                xirrFunction.setEnable(false);
            }

            FreeRefFunction[] functionImpls = { new GoalSeekFunction(workbook, eventDispatcher), xirrFunction };

            UDFFinder udfs = new DefaultUDFFinder(functionNames, functionImpls);
            UDFFinder udfToolpack = new AggregatingUDFFinder(udfs);

            workbook.addToolPack(udfToolpack);

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            formulaEvaluator.clearAllCachedResultValues();
            LocalDateTime start = LocalDateTime.now();
            System.out.println("=========" + start);
            formulaEvaluator.evaluateAll();
            LocalDateTime end = LocalDateTime.now();
            System.out.println("=========" + end);

            // Konvertálás azonnali típusba alapértelmezett időzóna eltolással
            Instant startInstant = start.toInstant(ZoneOffset.UTC);
            Instant endInstant = end.toInstant(ZoneOffset.UTC);

            // Időtartam kiszámítása
            long millis = Duration.between(startInstant, endInstant).toMillis();
            System.out.println(millis + " ms");


            try (FileOutputStream fos = new FileOutputStream(outFilePath)) {
                workbook.write(fos);
            }
        }
    }
}
