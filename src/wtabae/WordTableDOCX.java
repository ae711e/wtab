package wtabae;

/**
 * Created by ������� on 03.07.2016.
 * (C) 2016, Aleksey Eremin
 *
 * ��������� ������� � �����  .docx
*/
// Modify:
// 2016-07-03 ������� ������� "�������� ������" emptedStarngeCell()
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

// ���������� ������� �� ����������� � ����� DOCX
public class WordTableDOCX extends myParser {
    @Override
    boolean parse(String inputFileName, String outputFileName)
    {
        final String strNEED_TEXT="������������ ����� ���������.";   	// �������� ��������� �������
        String str;
        int cn, cnt;
        int i, n;
        boolean result=false;
        String DiRi;		// �������� �������� � ����������
        File file;
        FileInputStream fileInputStream=null;
        XWPFDocument docx;
        try {
            file=new File(inputFileName);
            fileInputStream = new FileInputStream(file);
            // ��������� ���� � ��������� ��� ���������� � ������ XWPFDocument
            docx = new XWPFDocument(OPCPackage.open(fileInputStream));  // word 2007
            // ������� ��� �������� �������� �����
            DiRi=file.getParent();  // �������� �������� � ����������
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
        XWPFTable mytable = null;  // ��������� ������� �� ������� ����������
        int mycol=0; // ������� � ������� ������
        // ��������� �� �������� ��������� � ������� ������ �������,
        // �� ���� �����, � ������� � ��������� ������� ��������� �������� needText
        tableIter = docx.getTablesIterator();
        while(tableIter.hasNext()) {
            table = tableIter.next();		// ��������� �������
            //printTable(table);
            cellList=table.getRow(0).getTableCells(); // ����� ����� ������ ������
            // ���� ������� ����������
            cn=cellList.size()-1;  // ������ ��������� ������ (�������)
            // ������� � 3 ��������� (������ ��������� ������� 2) ��������
            // �� �������� ������ � ������� ��
            if(cn == 2) {
                emptedStrangeCell(table);
            }
            // ������� � 5 ��������� �������� �� ���������� ��������� �������
            // ���� ������� �� �����������
            if (cn == 4) {
                str=cellList.get(cn).getText();  // ��������� ������� (��� ������ ���� ����� ������)
                if( str.regionMatches(0, strNEED_TEXT, 0, strNEED_TEXT.length()) ) { // ������� � ������� ������� �������� �������
                    mytable=table;	// �������� �������
                    mycol=cn;	// �������� ������ �������
                    break;
                }
            }
        } // end while
        //
        // ����� � ��������  - ����� �� �������?
        // �������, ���� ��������� ������� ������ 0, ������ �����
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
        // ������ �� ��������� ������� � ����� � ����� ������� ������ � ������� ��� ��������
        cnt=0;
        rowList = mytable.getRows();	// ������ ����� �������
        n=rowList.size();				// ���-�� ����� �������
        // ������ ��������� ������ ������ � ����� �������,
        // � ���� ����� ��� - ������� ������
        int kolnar=0;
        for(i=n-1; i>0; i--) {
            row=rowList.get(i);					// ����� ������ � �������� i
            cellList = row.getTableCells();		// �� ������ �������� ������ ����� (�������)
            str=cellList.get(mycol).getText();	// ��������� ������� - ��� ����� � ���������
            // ���� ��������� (����������) ����� ����� .jpg
            if(str.endsWith(".jpg")) {  // str.regionMatches(l-4, ".jpg",0,4) l=str.length(); // ����� ������
                // �������� ��� �����
                file=new File(DiRi,str);
                kolnar++;   // ���-�� ��������� (����������� ������)
                // �������� ���� �� ����� ����?
                if(!file.exists()) {
                    // ����� ���
                    mytable.removeRow(i);		// ������ ������ i �� �������
                    cnt++;
                    ////Log("Not exist " + cnt + ") " + str);
                }  // if (!file.exists())
            } // end if jpg
            //
        }
        //
        Log("Delete rows: " + cnt);
        //        
        // ���� ���� ��������� (�������� �������������� ����������)
        if(cnt>0) {
        	kolnar=0;  // ���-�� ��������� � �������, ������ ������� ������)
            // ������������ ������ �������
            rowList = mytable.getRows();	// ������ ����� �������
            n=rowList.size();				// ���-�� ����� �������
            cnt=0;
            // ������ ��������� ������� �� ������ ������ i=1 (� ������-�� ���������) � ������� ������
            for(i=1; i<n; i++) {
                row=rowList.get(i);					// ����� ������ � �������� i
                cellList = row.getTableCells();		// �� ������ �������� ������ ����� (�������)
                cell=cellList.get(0);				// ����� ������ ������ (������ �������, ������� ������)
                removeParagraphs(cell);				// ������� ��������� ��������� � ������
                str=Integer.toString(i);
                cell.setText(str); 					// ������� �����
                kolnar++;   // ���-�� ���������
            } // for
            //
            // ������� ���������� �������� � �������� ����
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
        // ������� ���-�� ��������� � ����� ������
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

    // ������� ��������� �� ��������� ������
    private static void removeParagraphs(XWPFTableCell tableCell)
    {
        // ������� � ������!
        for(int i = tableCell.getParagraphs().size()-1; i >=0; i--) {
            tableCell.removeParagraph(i);
        }
    }

    // ����������� �������
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

    // �������� �������� ������ �������
    private static void emptedStrangeCell(XWPFTable table)
    {
        final String str_STRANGE_TEXT="tbllkhdrcols";    // �������� �����
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
                    removeParagraphs(cell);                // ������� ��������� ��������� � ������
                    cell.setText(" "); 					// ������� ������
                    ////System.out.println("Strange text empted");
                }
            } // end while cell
            //
        } // end while row
        //
    } // end emptedStrangeCell()

    
    
}  // end WordTableDOCX(

