package wtabae;

/**
 * Created by Алексей on 03.07.2016.
 * (C) 2016, Aleksey Eremin
 *
 * обработка таблицы в файле  .docx
*/
// Modify:
// 2016-07-03 добавил очистку "странной ячейки" emptedStarngeCell()
//

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

// переделать таблицу со скриншотами в файле DOCX
public class WordTableDOCX extends myParser {
    @Override
    boolean parse(String inputFileName, String outputFileName)
    {
        final String strNEED_TEXT="Наименование файла скриншота.";   	// название требуемой колонки
        String str;
        int cn, cnt;
        int i, n;
        boolean result=false;
        String DiRi;		// название каталога с картинками
        File file;
        FileInputStream fileInputStream=null;
        XWPFDocument docx;
        try {
            file=new File(inputFileName);
            fileInputStream = new FileInputStream(file);
            // открываем файл и считываем его содержимое в объект XWPFDocument
            docx = new XWPFDocument(OPCPackage.open(fileInputStream));  // word 2007
            // Получим имя каталога входного файла
            DiRi=file.getParent();  // название каталога с картинками
            //
            fileInputStream.close();
        } catch(IOException | InvalidFormatException ex) {
            ex.printStackTrace();
            return false;
        }
        //
        Iterator<XWPFTable> tableIter;
        List<XWPFTableRow> rowList;
        List<XWPFTableCell> cellList;
        XWPFTable table;
        XWPFTableRow row;
        XWPFTableCell cell;
        XWPFTable mytable = null;  // требуемая таблица со списокм скриншотов
        int mycol=0; // колонка с именами файлов
        // пройдемся по таблицам документа в поисках нужной таблицы,
        // то есть такой, у которой в последнем столбце заголовка написано needText
        tableIter = docx.getTablesIterator();
        while(tableIter.hasNext()) {
            table = tableIter.next();		// очередная таблица
            //printTable(table);
            cellList=table.getRow(0).getTableCells(); // набор ячеек первой строки
            // ищем таблицу скриншотов
            cn=cellList.size()-1;  // индекс последней ячейки (столбца)
            // таблицы с 3 колонками (индекс последней колонки 2) проверим
            // на странную ячейку и очистим ее
            if(cn == 2) {
                emptedStrangeCell(table);
            }
            // таблицы с 5 колонками проверим на содержимое последней колонки
            // ищем таблицу со скриншотами
            if (cn == 4) {
                str=cellList.get(cn).getText();  // последняя колонка (там должны быть имена файлов)
                if( str.regionMatches(0, strNEED_TEXT, 0, strNEED_TEXT.length()) ) { // сравним с искомым текстом заоловка столбца
                    mytable=table;	// запомним таблицу
                    mycol=cn;	// запомним индекс столбца
                    break;
                }
            }
        } // end while
        //
        // вышли и проверим  - нашли ли таблицу?
        // считаем, если последняя колонка больше 0, значит нашли
        if (mycol<1) {
            Log("?ERROR-Not found need table");
            try {
                docx.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
            return false;
        }
        //
        ////printTable(mytable);
        // пойдем по найденной таблице с конца и будем удалять строки у которых нет картинок
        cnt=0;
        rowList = mytable.getRows();	// список строк таблицы
        n=rowList.size();				// кол-во строк таблицы
        // начнем проверять список файлов с конца таблицы,
        // и если файла нет - удалять строку
        int kolnar=0;
        for(i=n-1; i>0; i--) {
            row=rowList.get(i);					// взять строку с индексом i
            cellList = row.getTableCells();		// из строки получить список ячеек (колонки)
            str=cellList.get(mycol).getText();	// последняя колонка - имя файла с картинкой
            // если окончание (расширение) имени файла .jpg
            if(str.endsWith(".jpg")) {  // str.regionMatches(l-4, ".jpg",0,4) l=str.length(); // длина строки
                // получили имя файла
                file=new File(DiRi,str);
                kolnar++;   // кол-во нарушений (изначальное сейчас)
                // проверим есть ли такой файл?
                if(!file.exists()) {
                    // файла нет
                    mytable.removeRow(i);		// удалим строку i из таблицы
                    cnt++;
                    ////Log("Not exist " + cnt + ") " + str);
                }  // if (!file.exists())
            } // end if jpg
            //
        }
        //
        Log("Delete rows: " + cnt);
        //        
        // если было изменение (удаление несуществующих скриншотов)
        if(cnt>0) {
        	kolnar=0;  // кол-во нарушений в таблице, начнем считать заново)
            // переномеруем строки таблицы
            rowList = mytable.getRows();	// список строк таблицы
            n=rowList.size();				// кол-во строк таблицы
            cnt=0;
            // начнем проходить таблицу со второй строки i=1 (в первой-то заголовок) и ставить номера
            for(i=1; i<n; i++) {
                row=rowList.get(i);					// взять строку с индексом i
                cellList = row.getTableCells();		// из строки получить список ячеек (колонки)
                cell=cellList.get(0);				// взять первую ячейка (первый столбец, текущая строка)
                removeParagraphs(cell);				// удалить имеющиеся параграфы в ячейке
                str=Integer.toString(i);
                cell.setText(str); 					// вписать номер
                kolnar++;   // кол-вл нарушений
            } // for
            //
            // запишем измененный документ в выходной файл
            try {
                FileOutputStream out = new FileOutputStream(outputFileName);
                docx.write(out);
                out.close();
                ////Log("Wrote file: " + outputFileName);
                result=true;
            } catch (FileNotFoundException ex) {
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
            //
        }
        //
        Log("--------------------------");
        Log("KOL-VO NARUSHENII: " + kolnar);
        Log("--------------------------");
        // загоним кол-во нарушений в буфер обмена
        TextTransfer clp = new TextTransfer();
        clp.setData(String.valueOf(kolnar));
        //
        //
        try {
            docx.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        //
        return result;
    } // parse()

    // удалить параграфы из табличной ячейки
    private static void removeParagraphs(XWPFTableCell tableCell)
    {
        // удаляем с хвоста!
        for(int i = tableCell.getParagraphs().size()-1; i >=0; i--) {
            tableCell.removeParagraph(i);
        }
    }

    // распечатать таблицу
    private static void printTable(XWPFTable table)
    {
        List<XWPFTableRow> rowList;
        Iterator<XWPFTableRow> rowIter;
        List<XWPFTableCell> cellList;
        Iterator<XWPFTableCell> cellIter;
        XWPFTableRow row;
        XWPFTableCell cell;
        System.out.println("[Begin table]");
        rowList = table.getRows();
        rowIter = rowList.iterator();
        while(rowIter.hasNext()) {
            row = rowIter.next();
            cellList = row.getTableCells();
            cellIter = cellList.iterator();
            while(cellIter.hasNext()) {
                cell = cellIter.next();
                System.out.print("|" + cell.getText());
            }
            System.out.println("|");
        }
        System.out.println("[End table]");
    }

    // очистить странную ячейку таблицу
    private static void emptedStrangeCell(XWPFTable table)
    {
        final String str_STRANGE_TEXT="tbllkhdrcols";    // странный текст
        String str;
        List<XWPFTableRow> rowList;
        Iterator<XWPFTableRow> rowIter;
        List<XWPFTableCell> cellList;
        Iterator<XWPFTableCell> cellIter;
        XWPFTableRow row;
        XWPFTableCell cell;
        //
        rowList = table.getRows();
        rowIter = rowList.iterator();
        while(rowIter.hasNext()) {
            row = rowIter.next();
            cellList = row.getTableCells();
            cellIter = cellList.iterator();
            while(cellIter.hasNext()) {
                cell = cellIter.next();
                str=cell.getText();
                if (str.compareTo(str_STRANGE_TEXT)==0) {  // str.regionMatches(0, str_STRANGE_TEXT,0,str_STRANGE_TEXT.length())
                    removeParagraphs(cell);                // удалить имеющиеся параграфы в ячейке
                    cell.setText(" "); 					// вписать пробел
                    ////System.out.println("Strange text empted");
                }
            } // end while cell
            //
        } // end while row
        //
    } // end emptedStrangeCell()

    
    
}  // end WordTableDOCX(

