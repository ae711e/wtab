package wtabae;

/*
 Работа с буфером обмена
 (C) 2016, Aleksey Eremin
 по мотивам http://blog.dimka3210.ru/2013/04/java-java-clipboard-example_12.html
 */

import javax.tools.Tool;
import java.awt.*;
import java.awt.datatransfer.*;
import java.io.IOException;

public class TextTransfer implements ClipboardOwner {
	StringSelection stringSelection;

	@Override
	public void lostOwnership(Clipboard clipboard, Transferable contents) {
		// TODO Auto-generated method stub
	}
	
	public void setData(String data)
	{
	    stringSelection = new StringSelection(data);
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents(stringSelection, this);
	}

        public String getData()
        {
            String str;
            try {
                Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                str = (String) clipboard.getData(DataFlavor.stringFlavor);
            } catch (IOException|UnsupportedFlavorException e) {
                str="Empty clipboard";
            }
            return str;
        }

}
