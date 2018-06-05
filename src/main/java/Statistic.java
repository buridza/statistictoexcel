import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Statistic {


    public static void main(String[] args) {
        List<Long> move = new ArrayList<Long>();
        List<Long> countTime = new ArrayList<Long>();

        Statistic statistic = new Statistic();
        statistic.filling(move, countTime);

        XSSFWorkbook workbook = new XSSFWorkbook();
        workbook.createSheet("sheet for denis");

        statistic.toExcel(workbook, move, 0);
        statistic.toExcel(workbook, countTime, 1);


    }

    Long size(final List<Long> list, final int index) {
        long size;
        if (list.size() > index + 1000)
            size = index + 1000;
        else size = list.size();

        return size;
    }

    void toExcel(final XSSFWorkbook workbook, final List<Long> list, final int columnIndex) {

        int i = 0;
        boolean status = false;
        while (size(list, i)-1 < list.size() && !status) {
            XSSFSheet sheet;
            XSSFRow row;
            XSSFCell cell;
            for (; i < size(list, i) && i < 1048576; i++) {
                sheet = workbook.getSheet("sheet for denis");
                row = sheet.createRow(i);
                cell = row.createCell(columnIndex);
                cell.setCellValue(list.get(i));
            }
            try {
                workbook.write(new FileOutputStream("statistic.xlsx"));
                System.out.println("Таблица " + columnIndex + " готова");
                status=true;
            } catch (IOException io) {
                io.printStackTrace();
            }
        }
    }

    void filling(final List<Long> move, final List<Long> countTime) {
        System.out.println("запись начата");
        int x;
        int y;
        x = MouseInfo.getPointerInfo().getLocation().x;
        y = MouseInfo.getPointerInfo().getLocation().y;
        long startTime = System.currentTimeMillis();
        long timeMillis = startTime;
        while (timeMillis - startTime < 5000) {
            if ((x != MouseInfo.getPointerInfo().getLocation().x || y != MouseInfo.getPointerInfo().getLocation().y)) {
                move.add(System.currentTimeMillis() - timeMillis);
            } else if (System.currentTimeMillis()!=timeMillis) {
                countTime.add(System.currentTimeMillis() - timeMillis);
            }
            timeMillis = System.currentTimeMillis();
            x = MouseInfo.getPointerInfo().getLocation().x;
            y = MouseInfo.getPointerInfo().getLocation().y;
        }
        System.out.println("Остановлена");

    }


}
