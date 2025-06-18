package studio.intuitech.functions;

import java.io.FileInputStream;
import java.io.IOException;

public class CalculationExcelTemplateFactory {

    private static final byte[] NORMAL_TEMPLATE;
    private static final byte[] ONLY_XIRR_TEMPLATE;

    private static final String NORMAL_TEMPLATE_XLSX = "./normal.xlsx";
    private static final String ONLY_XIRR_TEMPLATE_XLSX = "./xirr.xlsx";

    static {
        try (FileInputStream fileInputStream = new FileInputStream(NORMAL_TEMPLATE_XLSX)) {
            NORMAL_TEMPLATE = fileInputStream.readAllBytes();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        try (
            FileInputStream fileInputStream = new FileInputStream(ONLY_XIRR_TEMPLATE_XLSX)) {
            ONLY_XIRR_TEMPLATE = fileInputStream.readAllBytes();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static byte[] getCalculationExcelTemplate(CalculationExcelType calculationExcelType) {
        switch (calculationExcelType) {
            case NORMAL:
                return NORMAL_TEMPLATE;
            case ONLY_XIRR:
                return ONLY_XIRR_TEMPLATE;
        }
        return new byte[0];
    }

}
