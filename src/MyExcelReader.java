import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;





/**
 * <ul>
 * <li>Title:[POI基础上的Excel数据读取工具]</li>
 * <li>Description: [支持Excell2003,Excell2007,自动格式化数值型数据,自动格式化日期型数据]</li>
 * <li>Copyright 2009 RoadWay Co., Ltd.</li>
 * <li>All right reserved.</li>
 * <li>Created by [惠万鹏] [Jan 20, 2010]</li>
 * <li>Midified by [modifier] [modified time]</li>
 * 
 * <li>所需Jar包列表</li>
 * <li>poi-3.6-20091214.jar</li>
 * <li>poi-contrib-3.6-20091214.jar</li>
 * <li>poi-examples-3.6-20091214.jar</li>
 * <li>poi-ooxml-3.6-20091214.jar</li>
 * <li>poi-ooxml-schemas-3.6-20091214.jar</li>
 * <li>poi-scratchpad-3.6-20091214.jar</li>
 * <li>xmlbeans-2.3.0.jar</li>
 * <ul>
 * 
 * @version 1.0
 */
class POIExcelUtil
{
    /** 总行数 */
    private int totalRows = 0;
    
    /** 总列数 */
    private int totalCells = 0;
    
    /** 构造方法 */
    public POIExcelUtil()
    {}
    
    /**
     * <ul>
     * <li>Description:[根据文件名读取excel文件]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @param fileName
     * @return
     * @throws Exception
     */
    public ArrayList<ArrayList<String>> read(String fileName)
    {
        ArrayList<ArrayList<String>> dataLst = new ArrayList<ArrayList<String>>();
        
        /** 检查文件名是否为空或者是否是Excel格式的文件 */
        if (fileName == null || !fileName.matches("^.+\\.(?i)((xls)|(xlsx))$"))
        {
            return dataLst;
        }
        
        boolean isExcel2003 = true;
        /** 对文件的合法性进行验证 */
        if (fileName.matches("^.+\\.(?i)(xlsx)$"))
        {
            isExcel2003 = false;
        }
        
        /** 检查文件是否存在 */
        File file = new File(fileName);
        if (file == null || !file.exists())
        {
            return dataLst;
        }
        
        try
        {
            /** 调用本类提供的根据流读取的方法 */
            dataLst = read(new FileInputStream(file), isExcel2003);
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
        
        /** 返回最后读取的结果 */
        return dataLst;
    }
    
    /**
     * <ul>
     * <li>Description:[根据流读取Excel文件]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @param inputStream
     * @param isExcel2003
     * @return
     */
    public ArrayList<ArrayList<String>> read(InputStream inputStream,
            boolean isExcel2003)
    {
        ArrayList<ArrayList<String>> dataLst = null;
        try
        {
            /** 根据版本选择创建Workbook的方式 */
            Workbook wb = isExcel2003 ? new HSSFWorkbook(inputStream):new HSSFWorkbook(inputStream);
            dataLst = read(wb);
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }
        return dataLst;
    }
    
    /**
     * <ul>
     * <li>Description:[得到总行数]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @return
     */
    public int getTotalRows()
    {
        return totalRows;
    }
    
    /**
     * <ul>
     * <li>Description:[得到总列数]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @return
     */
    public int getTotalCells()
    {
        return totalCells;
    }
    
    /**
     * <ul>
     * <li>Description:[读取数据]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @param wb
     * @return
     */
    private ArrayList<ArrayList<String>> read(Workbook wb)
    {
        ArrayList<ArrayList<String>> dataLst = new ArrayList<ArrayList<String>>();
        
        /** 得到第一个shell */
        Sheet sheet = wb.getSheetAt(5);
        this.totalRows = sheet.getPhysicalNumberOfRows();
        if (this.totalRows >= 1 && sheet.getRow(0) != null)
        {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        
        /** 循环Excel的行 */
        for (int r = 0; r <=this.totalRows; r++)
        {
            Row row = sheet.getRow(r);
            if (row == null)
            {
                continue;
            }
            
            ArrayList<String> rowLst = new ArrayList<String>();
            /** 循环Excel的列 */
            for (short c = 0; c < this.getTotalCells(); c++)
            {
                Cell cell = row.getCell(c);
                String cellValue = "";
                if (cell == null)
                {
                    rowLst.add(cellValue);
                    continue;
                }
                
                /** 处理数字型的,自动去零 */
                /** 在excel里,日期也是数字,在此要进行判断 */
               if (Cell.CELL_TYPE_NUMERIC == cell.getCellType())
                {
                        //cellValue = getRightStr(cell.getNumericCellValue() + "");
            	   cellValue=cell.getNumericCellValue()+"";
                }
                /** 处理字符串型 */
               else if (Cell.CELL_TYPE_STRING == cell.getCellType())
                {
                    cellValue = cell.getStringCellValue();
                }
                /** 处理布尔型 */
                else if (Cell.CELL_TYPE_BOOLEAN == cell.getCellType())
                {
                    cellValue = cell.getBooleanCellValue() + "";
                }
                /** 其它的,非以上几种数据类型 */
                else
                {
                    cellValue = cell.toString() + "";
                }
                
                rowLst.add(cellValue);
            }
            dataLst.add(rowLst);
        }
        return dataLst;
    }
    
    /**
     * <ul>
     * <li>Description:[正确地处理整数后自动加零的情况]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @param sNum
     * @return
     */
    private String getRightStr(String sNum)
    {
        DecimalFormat decimalFormat = new DecimalFormat("#.000000");
        String resultStr = decimalFormat.format(new Double(sNum));
        if (resultStr.matches("^[-+]?\\d+\\.[0]+$"))
        {
            resultStr = resultStr.substring(0, resultStr.indexOf("."));
        }
        return resultStr;
    }
    
    /**
     * <ul>
     * <li>Description:[测试main方法]</li>
     * <li>Created by [Huyvanpull] [Jan 20, 2010]</li>
     * <li>Midified by [modifier] [modified time]</li>
     * <ul>
     * 
     * @param args
     * @throws Exception
     */
    /*public static void main(String[] args) throws Exception
    {
        List<ArrayList<String>> dataLst = new POIExcelUtil()
                .read("e:/Book1_shao.xls");
        for (ArrayList<String> innerLst : dataLst)
        {
            StringBuffer rowData = new StringBuffer();
            for (String dataStr : innerLst)
            {
                rowData.append(",").append(dataStr);
            }
            if (rowData.length() > 0)
            {
                System.out.println(rowData.deleteCharAt(0).toString());
            }
        }
    }  */
    
    public void replaceExcel(String path1,String path2,ArrayList<ArrayList<String>> totalList){
    	try {
			POIFSFileSystem fs=new POIFSFileSystem(new FileInputStream(path2));
			HSSFWorkbook wb=new HSSFWorkbook(fs);
			HSSFSheet sheet=wb.getSheetAt(5);
			for(ArrayList<String> list:totalList){
				for(int i=1;i<=totalRows;i++){
					HSSFRow my_Row=sheet.getRow(i);
					HSSFCell cell1=my_Row.getCell((short)3);
					HSSFCell cell2=my_Row.getCell((short)4);
					//System.out.println(list.get(0));
					System.out.println(list);
					System.out.println(totalRows);
					String name=null;
					if(list.size()>0){
					     name=list.get(0);
					}
					if(name!=null){
						/*if(name.equals(cell1.getStringCellValue())){
							cell2.setCellValue(name);
							FileOutputStream fos=new FileOutputStream(path2);
							wb.write(fos);
							fos.close();
					    }*/
						System.out.println(my_Row.getPhysicalNumberOfCells());
					//String name2=my_Row.getCell((short)3).getStringCellValue();
					//if(name2!=null)
						//System.out.println(name2);
				  }
				}
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
    
    
}









/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Super.wang
 */
public class MyExcelReader extends javax.swing.JFrame implements ActionListener{

    /**
     * Creates new form MyExcelReader
     */
    public MyExcelReader() {
        initComponents();
    }


    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jTextField1 = new javax.swing.JTextField();
        jTextField2 = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox();
        jLabel3 = new javax.swing.JLabel();
        jButton3 = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jLabel1.setText("工资表：");

        jLabel2.setText("卡号表：");

        jTextField1.setText(" ");
        jTextField1.setColumns(20);
        jTextField1.setEditable(false);
        jTextField2.setText(" ");
        jTextField2.setColumns(20);
        jTextField2.setEditable(false);
        jTextField2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField2ActionPerformed(evt);
            }
        });

        jButton1.setText("浏览");

        jButton2.setText("浏览");

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "工资填充", "新功能1", "新功能2", "新功能3" }));

        jLabel3.setText("操作：");

        jButton3.setText("确定");

        jButton4.setText("取消");
        
        jButton1.addActionListener(this);
        jButton2.addActionListener(this);
        jButton3.addActionListener(this);
        jButton4.addActionListener(this);
        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, 229, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel2)
                            .addComponent(jLabel3))
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(16, 16, 16)
                                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 82, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE))
                            .addGroup(layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 80, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addComponent(jButton4, javax.swing.GroupLayout.PREFERRED_SIZE, 80, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addComponent(jTextField2))))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 67, Short.MAX_VALUE)
                    .addComponent(jButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(24, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(40, 40, 40)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1))
                .addGap(32, 32, 32)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jTextField2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton2))
                .addGap(28, 28, 28)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel3))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 58, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton3)
                    .addComponent(jButton4))
                .addGap(52, 52, 52))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jTextField2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField2ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField2ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
       /* try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MyExcelReader.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MyExcelReader.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MyExcelReader.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MyExcelReader.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MyExcelReader().setVisible(true);
            }
        });*/
    	 new MyExcelReader().setVisible(true);
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField2;
    private String path1=null,path2=null;
    // End of variables declaration//GEN-END:variables
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if(e.getSource()==jButton1){
		    JFileChooser chooser = new JFileChooser();
		    FileNameExtensionFilter filter = new FileNameExtensionFilter("MicroSoft Excel文档","xls");
		    chooser.setFileFilter(filter);
		    int returnVal = chooser.showOpenDialog(this);
		    if(returnVal == JFileChooser.APPROVE_OPTION) {
		            path1=chooser.getSelectedFile().getAbsolutePath();
		            jTextField1.setText(path1);
		    }

		}
		if(e.getSource()==jButton2){
		    JFileChooser chooser = new JFileChooser();
		    FileNameExtensionFilter filter = new FileNameExtensionFilter("MicroSoft Excel文档","xls");
		    chooser.setFileFilter(filter);
		    int returnVal = chooser.showOpenDialog(this);
		    if(returnVal == JFileChooser.APPROVE_OPTION) {
		            path2=chooser.getSelectedFile().getAbsolutePath();
		            jTextField2.setText(path2);
		    }
		}
		if(e.getSource()==jButton3){
			if((path1==null)||(path2==null)){
				JOptionPane.showMessageDialog(this, "请选择文件!");
			} else{
				POIExcelUtil poExl=new POIExcelUtil();
				ArrayList<ArrayList<String>> totalList=new ArrayList<ArrayList<String>>();
				totalList=poExl.read(path1);
				poExl.replaceExcel(path1, path2, totalList);
			}
		}
		if(e.getSource()==jButton4){
			System.exit(0);
		}
	}
}
