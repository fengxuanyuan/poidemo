package com.fxy;
/*
 * @Description :
 * @Author fengxuanyuan2010@foxmail.com
 * @Date 2020/4/20 16:28
 */

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class PoiServlet extends HttpServlet {


    private static final long serialVersionUID = 1L;

    @Override
    public void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
            doPost(request,response);

    }

    @Override
    public void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        String method = req.getParameter("_m");
        if("poi_down".equals(method)){
            poi_down(req,resp);
        }else if("poi_upload".equals(method)){
            poi_upload(req,resp);
        }
    }


    private void poi_upload(HttpServletRequest request,
                            HttpServletResponse response) {
        if(ServletFileUpload.isMultipartContent(request)){
            DiskFileItemFactory factory = new DiskFileItemFactory();
            factory.setSizeThreshold(1024*512);
            factory.setRepository(new File("D:/tempload"));
            ServletFileUpload fileUpload=new ServletFileUpload(factory);
            fileUpload.setFileSizeMax(10*1024*1024);//设置最大文件大小
            try {
                @SuppressWarnings("unchecked")
                List<FileItem> items=fileUpload.parseRequest(request);//获取所有表单
                for(FileItem item:items){
                    //判断当前的表单控件是否是一个普通控件
                    if(!item.isFormField()){
                        //是一个文件控件时
                        String excelFileName = new String(item.getName().getBytes(), "utf-8"); //获取上传文件的名称
                        //上传文件必须为excel类型,根据后缀判断(xls)
                        String excelContentType = excelFileName.substring(excelFileName.lastIndexOf(".")); //获取上传文件的类型
                        System.out.println("上传文件名:"+excelFileName);
                        System.out.println("文件大小:"+item.getSize());
                        System.out.println("\n---------------------------------------");
                        if(".xlsx".equalsIgnoreCase(excelContentType)){
                           // POIFSFileSystem fileSystem = new XSSFWorkbook(item.getInputStream());
                            XSSFWorkbook workbook = new XSSFWorkbook(item.getInputStream());
                            XSSFSheet sheet = workbook.getSheetAt(0);
                            int rows = sheet.getPhysicalNumberOfRows();
                            for (int i = 2; i < rows; i++) {
                                XSSFRow row = sheet.getRow(i);

                                    int columns = row.getPhysicalNumberOfCells();

                                    for (int j = 0; j < columns; j++) {
                                        XSSFCell cell = row.getCell(j);
                                        if(cell != null && cell.getCellType() != BLANK) {
                                            cell.setCellType(STRING);
                                            String value = this.getCellStringValue(cell);
                                            System.out.print(value + "|");
                                        }
                                    }

                                System.out.println("\n---------------------------------------");
                            }
                            System.out.println("success！");
                        }else if(".xls".equalsIgnoreCase(excelContentType)){
                            POIFSFileSystem fileSystem = new POIFSFileSystem(item.getInputStream());
                            HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
                            HSSFSheet sheet = workbook.getSheetAt(0);

                            int rows = sheet.getPhysicalNumberOfRows();
                            for (int i = 0; i < rows; i++) {
                                HSSFRow row = sheet.getRow(i);
                                if(row.getFirstCellNum()>-1) {
                                    int columns = row.getPhysicalNumberOfCells();
                                    for (int j = 0; j < columns; j++) {
                                        HSSFCell cell = row.getCell(j);
                                        String value = this.getCellStringValue(cell);
                                        System.out.print(value + "|");
                                    }
                                }
                                System.out.println("\n---------------------------------------");
                            }
                            System.out.println("success！");
                        }else{
                            System.out.println("必须为excel类型");
                        }
                        //顺便把文件保存到硬盘,防止重名
//                        String newName=new SimpleDateFormat("yyyyMMDDHHmmssms").format(new Date());
//                        File file = new File("d:/upload");
//                        if(!file.exists()){
//                            file.mkdir();
//                        }
//                        item.write(new File("d:/upload/"+newName+excelContentType));
                        response.sendRedirect("index.jsp");
                    }
                }
            }catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private void poi_down(HttpServletRequest request,
                          HttpServletResponse response) {
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=data.xls");

        ServletOutputStream stream = null;
        try {
            stream = response.getOutputStream();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//左右居中样式

        HSSFSheet sheet = workbook.createSheet("我的联系人");
        sheet.setColumnWidth(0, 2000);
        sheet.setColumnWidth(1, 5000);
        //创建表头(第一行)
        HSSFRow row = sheet.createRow(0);
        //列
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("姓名");
        cell.setCellStyle(style);
        HSSFCell cell2 = row.createCell(1);
        cell2.setCellValue("电话");
        cell2.setCellStyle(style);

        //创建数据行
        for(int i =1;i<=20;i++) {
            HSSFRow newrow = sheet.createRow(i);
            newrow.createCell(0).setCellValue("tom"+i);
            newrow.createCell(1).setCellValue("135816****"+i);
        }
        try {
            workbook.write(stream);
            System.out.println("下载成功");
            stream.flush();
            stream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    //获取单元格内不同类型值
    public String getCellStringValue(Cell cell) {
        try {
            String cellValue = "";
            switch (cell.getCellType()) {
                case STRING:
                    cellValue = cell.getStringCellValue();
                    if(cellValue.trim().equals("")||cellValue.trim().length()<=0)
                        cellValue=" ";
                    break;
                case NUMERIC:
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case FORMULA:
                    cell.setCellType(CellType.NUMERIC);
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                case BLANK:
                    cellValue=" ";
                    break;
                case BOOLEAN:
                    break;
                case ERROR:
                    break;
                default:
                    break;
            }
            return cellValue.trim();
        } catch (Exception e) {
            return null;
        }
    }

}
