import javax.swing.*;
import java.awt.*;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
/**
 * @author 93037
 */
public class ExcelUtil extends JFrame {
    JLabel label1 =new JLabel();
     JLabel label2 =new JLabel();
     JLabel label3 =new JLabel("第几页");
     JLabel label4 =new JLabel("标题行数");
     JLabel label5 =new JLabel("总表拆分输入地址");
     JLabel label6 =new JLabel("总表拆的列");
     JButton btnConfim=new JButton("确认");
     JButton btnSplit=new JButton("确认");
     JTextField path=new JTextField();
     JTextField search =new JTextField();
     JTextField pathForSplit =new JTextField();
     JTextField page =new JTextField();
     JTextField title =new JTextField();
     JTextField column =new JTextField();
    public static void main(String[] args) {
        setActionListener();
    }
    private static void setActionListener(){
        ExcelUtil excelutil = new ExcelUtil();
        excelutil.btnConfim.addActionListener(e -> {
            String pathStr =excelutil.path.getText().trim()+".xlsx";
            String searchStr =excelutil.search.getText().trim();
            int page=1;
            int title=1;
            if(!("".equals(excelutil.page.getText()))){
                page=Integer.parseInt(excelutil.page.getText().trim());
            }
            if(!("".equals(excelutil.title.getText()))){
                title=Integer.parseInt(excelutil.title.getText().trim());
            }
            excelutil.searchExcel(pathStr, searchStr,page,title);
            new MyDialog1(excelutil,"提取完成").setVisible(true);
        });
        excelutil.btnSplit.addActionListener(e -> {
            String pathStr =excelutil.pathForSplit.getText().trim()+".xlsx";
            int page=1;
            int title=1;
            int column=1;
            if(!("".equals(excelutil.page.getText()))){
                page=Integer.parseInt(excelutil.page.getText().trim());
            }
            if(!("".equals(excelutil.title.getText()))){
                title=Integer.parseInt(excelutil.title.getText().trim());
            }
            if(!("".equals(excelutil.title.getText()))){
                column=Integer.parseInt(excelutil.column.getText().trim());
            }
            Set<String> infos = excelutil.getInfo(pathStr,page,column,title);
            while (infos.iterator().hasNext()){
                Iterator<String> iterator=infos.iterator();
                excelutil.searchExcel(pathStr,iterator.next(),page,title,column);
                iterator.remove();
            }
            new MyDialog1(excelutil,"提取完成").setVisible(true);
        });
    }
    private ExcelUtil(){
        this.setVisible(true);
        setSize(800,100);
        this.setResizable(false);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        this.setTitle("EXCEL提取器");
        label1.setText("文件地址");
        label2.setText("索引信息");
        JPanel panel = new JPanel();
        panel.setLayout(new FlowLayout());
        path.setPreferredSize(new Dimension (200,30));
        search.setPreferredSize(new Dimension (200,30));
        page.setPreferredSize(new Dimension (30,30));
        pathForSplit.setPreferredSize(new Dimension(200,30));
        title.setPreferredSize(new Dimension(30,30));
        column.setPreferredSize(new Dimension(30,30));
        this.setLocation(500,400);
        this.add(panel);
        panel.add(label1);
        panel.add(path);
        panel.add(label2);
        panel.add(search);
        panel.add(label3);
        panel.add(page);
        panel.add(label4);
        panel.add(title);
        panel.add(btnConfim);
        panel.add(label5);
        panel.add(pathForSplit);
        panel.add(label6);
        panel.add(column);
        panel.add(btnSplit);
    }
    private void searchExcel(String path, String search,int page,int title){

        try {
            int rowNum = 0;
//  需要解析的Excel文件
            File file = new File(path);
//            创建Excel，读取文件内容
            XSSFWorkbook readWorkbook = new XSSFWorkbook(FileUtils.openInputStream(file));
//            提取的工作空间
            XSSFWorkbook writeWorkbook = new XSSFWorkbook();
            XSSFSheet writeSheet = writeWorkbook.createSheet();
//            读取默认第一个工作表Sheet
            XSSFSheet sheet = readWorkbook.getSheetAt(page-1);
            //标题转录
            for (int i = 0; i < title; i++) {
                copyRow(writeSheet.createRow(i),sheet.getRow(i));
                rowNum++;
            }
//            获取sheet中最后一行行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = title; i <= lastRowNum; i++) {
                //标题行结束后开始遍历
                XSSFRow row = sheet.getRow(i);
//                获取当前行最后单元格列号
                int lastCellNum = row.getLastCellNum();
                for (int j = 0; j < lastCellNum; j++) {
                    XSSFCell cell = row.getCell(j);
                    rowNum = getRowNum(search, rowNum, writeSheet, row, cell);
                }
            }
            outputName(search, writeWorkbook);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private int getRowNum(String search, int rowNum, XSSFSheet writeSheet, XSSFRow row, XSSFCell cell) {
        if(!(cell==null||"".equals(cell)||cell.getCellType() ==XSSFCell.CELL_TYPE_BLANK)){
            cell.setCellType(Cell.CELL_TYPE_STRING);
            //每一格的元素
            String value = cell.getStringCellValue();
            //目标字段匹配则转出此行所有信息
            if (search.equals(value)) {
                XSSFRow writeRow = writeSheet.createRow(rowNum);
                copyRow(writeRow,row);
                ++rowNum;
            }
        }
        return rowNum;
    }
    private void searchExcel(String path, String search,int page,int title,int column){

        try {
            int rowNum = 0;
//  需要解析的Excel文件
            File file = new File(path);
//            创建Excel，读取文件内容
            XSSFWorkbook readWorkbook = new XSSFWorkbook(FileUtils.openInputStream(file));
//            提取的工作空间
            XSSFWorkbook writeWorkbook = new XSSFWorkbook();
            XSSFSheet writeSheet = writeWorkbook.createSheet();
//            读取默认第一个工作表Sheet
            XSSFSheet sheet = readWorkbook.getSheetAt(page-1);
            //标题转录
            for (int i = 0; i < title; i++) {
                copyRow(writeSheet.createRow(i),sheet.getRow(i));
                rowNum++;
            }
//            获取sheet中最后一行行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = title; i <= lastRowNum; i++) {
                //标题行结束后开始遍历
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell = row.getCell(column-1);
                rowNum = getRowNum(search, rowNum, writeSheet, row, cell);
            }
            outputName(search, writeWorkbook);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private void outputName(String search, XSSFWorkbook writeWorkbook) {
        File newFile = new File("D:/"+search.replaceAll("[^\\u4e00-\\u9fa5]", "")+".xlsx");
        try {
            newFile.createNewFile();
//            将Excel内容存盘
            FileOutputStream stream = FileUtils.openOutputStream(newFile);
            writeWorkbook.write(stream);
            stream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private Set<String> getInfo(String path,int page,int column,int title) {
        //保存信息的列名
        Set<String> infos = new HashSet();
        //  需要解析的Excel文件
        File file = new File(path);
        try {
//            创建Excel，读取文件内容
            XSSFWorkbook readWorkbook = new XSSFWorkbook(FileUtils.openInputStream(file));
//            读取默认第一个工作表Sheet
            XSSFSheet sheet = readWorkbook.getSheetAt(page - 1);
//            获取sheet中最后一行行号
            int lastRowNum = sheet.getLastRowNum();
            for (int i = title; i <= lastRowNum; i++) {
                XSSFRow row = sheet.getRow(i);
                XSSFCell cell=row.getCell(column-1);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                if (!("".equals(cell.getStringCellValue()))){
                    //保存索引信息
                    infos.add(cell.getStringCellValue());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return infos;
    }
    private  void copyRow(XSSFRow newRow, XSSFRow oldRow){
        for (int i = 0; i < oldRow.getLastCellNum(); i++) {
            XSSFCell oldCell = oldRow.getCell(i);
            oldCell.setCellType(Cell.CELL_TYPE_STRING);
            XSSFCell newCell=newRow.createCell(i);
            newCell.setCellValue(oldCell.getStringCellValue());
        }
    }
    static class MyDialog1 extends JDialog {
        private static final long serialVersionUID = 1L;
        MyDialog1(JFrame frame, String str) {
            super(frame, "提示信息");
            Container conn = getContentPane();
            conn.add(new JLabel(str));
            setBounds(100, 100, 150, 150);
            setLocation(800,500);
        }
    } //提示窗口
}