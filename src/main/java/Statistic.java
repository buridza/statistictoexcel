import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

public class Statistic {


    private static final int MAX_ITERATIONS = 5000;

    String FILE = "str";

    public static void main(String[] args) {
        List<Long> move = new ArrayList<Long>(MAX_ITERATIONS);
        List<Long> countTime = new LinkedList<>();

        String str = (LocalDateTime.now().toString().replaceAll(":",".")+ " statistic.xlsx" );

        Statistic statistic = new Statistic();
        statistic.FILE = str;
        statistic.filling(move, countTime);

        //filtering works much faster with LinkedList rather than ArrayList
        Long previous = null;
        Iterator<Long> countInterator = countTime.iterator();
        while (countInterator.hasNext()) {
            Long current = countInterator.next();
            if (current.equals(previous)) {
                countInterator.remove();
            }
            previous = current;
        }

        //methods below use List.get(index) a lot, so convert LinkedList to ArrayList
        countTime = new ArrayList<>(countTime);

        //geenrate excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        statistic.toExcel(workbook, move, 0);
        statistic.toExcel(workbook, countTime, 1);
    }



    Long size(final List<Long> list, final int index) {
        return (long)((list.size() > index + 1000)?(index + 1000): list.size());
    }

    void toExcel(final XSSFWorkbook workbook, final List<Long> list, final int columnIndex) {

        int i = 0;
        boolean status = false;
        XSSFSheet sheet;
        XSSFRow row;
        XSSFCell cell;
        sheet = workbook.createSheet("sheet for denis" + columnIndex);
        while (size(list, i) - 1 < list.size() && !status) {

            for (; i < size(list, i) && i < 1048576; i++) {
                sheet = workbook.getSheet("sheet for denis" + columnIndex);
                row = sheet.createRow(i);
                cell = row.createCell(columnIndex);
                cell.setCellValue(list.get(i));
            }
            try {
                workbook.write(new FileOutputStream(FILE));
                System.out.println("Таблица " + columnIndex + " готова");
                status = true;
            } catch (IOException io) {
                io.printStackTrace();
            }
        }
    }

    void filling(final List<Long> move, final List<Long> countTime) {
        System.out.println("запись начата");
        int x = MouseInfo.getPointerInfo().getLocation().x;
        int y = MouseInfo.getPointerInfo().getLocation().y;

        long startTime = System.currentTimeMillis();
        long timeMillis = startTime;

        while (timeMillis - startTime < MAX_ITERATIONS) {
            if ((x != MouseInfo.getPointerInfo().getLocation().x || y != MouseInfo.getPointerInfo().getLocation().y)) {
                move.add(System.currentTimeMillis() - timeMillis);
                timeMillis = System.currentTimeMillis();
            } else {
                countTime.add(System.currentTimeMillis() - timeMillis);
            }

            x = MouseInfo.getPointerInfo().getLocation().x;
            y = MouseInfo.getPointerInfo().getLocation().y;
        }
        System.out.println("Остановлена");

    }


}
