import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * @author: hhhh
 * @createTime: 2022/2/20 20:36
 * @description:
 */
public class calculateAdress {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        //定义表格文件的位置
        File file = new File("test1.xlsx");
        //获取excel文件的输入流
        FileInputStream fileInputStream = new FileInputStream(file);//
        XSSFWorkbook xssfWorkbook = null;
        xssfWorkbook = new XSSFWorkbook(fileInputStream); //获取excel文件
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);//获取excel文件的第一个子表格
        int rowNum = sheet.getPhysicalNumberOfRows();//获得总行数
        HashMap<String, String> map = new HashMap<>();
        for (int i = 1; i < rowNum; i++) { //从表格的第二行开始遍历,一直到最后一行
            String start = sheet.getRow(i).getCell(0).toString();//得到第i+1行,第1列
            String end = sheet.getRow(i).getCell(1).toString();//得到第i+1行,第2列
            XSSFCell yanma = sheet.getRow(i).getCell(2);//得到第i+1行,第3列
            XSSFCell yanma2 = sheet.getRow(i).getCell(3);//得到第i+1行,第4列
            String mask = getMask(start, end);//调用函数得到子网掩码ip形式
            int intmask = mask2len(mask);//得到子网掩码位数形式
            String smask = String.valueOf(intmask);
            if (null != yanma || Objects.requireNonNull(yanma).toString() != null) {
                yanma.setCellValue("");//判空,如果为空就重置
            }
            yanma.setCellValue(start + "/" + smask);//写入位数形式的子网掩码

            if (null != yanma2 || Objects.requireNonNull(yanma).toString() != null) {
                yanma2.setCellValue("");//判空,如果为空就重置
            }
            yanma2.setCellValue(mask);//写入ip形式的子网掩码
        }
        FileOutputStream fos = null;
        fos = new FileOutputStream(file);//io流写入excel文件的对象
        xssfWorkbook.write(fos);//写入
        fos.close();//关闭io流
        xssfWorkbook.close();//关闭源文件对象
//        OutputStream out = null;
//       File test2 = new File("D:\\test1\\test1(计算后).xlsx");
//        try {
//            out = new FileOutputStream(test2);
//            xssfWorkbook.write(out);
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        } finally {
//            assert out != null;
//            out.close();
//            xssfWorkbook.close();
//        }
        System.out.println("感谢使用,!!!!操作已经全部完成");
    }

    private static byte pos[] = new byte[]{(byte) 128, 64, 32, 16, 8, 4, 2, 1};//初始化静态数组

    public static String getMask(String startIP, String endIP) {//获取子网掩码ip形式的方法
        byte start[] = getAddress(startIP);
        byte end[] = getAddress(endIP);
        byte mask[] = new byte[start.length];
        boolean flag = false;
        for (int i = 0; i < start.length; i++) {
            mask[i] = (byte) ~(start[i] ^ end[i]);
            if (flag) mask[i] = 0;
            if (mask[i] != -1) {
                mask[i] = getMask(mask[i]);
                flag = true;
            }
        }
        return toString(mask);
    }

    private static byte getMask(byte b) {//处理每一段的byte数值
        if (b == 0) return b;
        byte p = pos[0];
        for (int i = 0; i < 8; i++) {
            if ((b & pos[i]) == 0) break;
            p = (byte) (p >> 1);
        }
        p = (byte) (p << 1);
        return (byte) (b & p);
    }

    public static String toString(byte[] address) { //格式化子网掩码,转化为ip形式的字符串
        StringWriter sw = new StringWriter(16);
        sw.write(Integer.toString(address[0] & 0xFF));
        for (int i = 1; i < address.length; i++) {
            sw.write(".");
            sw.write(Integer.toString(address[i] & 0xFF));
        }
        return sw.toString();
    }

    public static int mask2len(String mask) {//将子网掩码ip形式 转化为 位数形式
        String[] masks = mask.split("\\.");
        //System.out.println(Arrays.toString(masks));
        int intmask =
                ((Integer.parseInt(masks[0]) & 0xFF) << 24) +
                        ((Integer.parseInt(masks[1]) & 0xFF) << 16) +
                        ((Integer.parseInt(masks[2]) & 0xFF) << 8) +
                        ((Integer.parseInt(masks[3]) & 0xFF));
        //System.out.println(intmask);
        int len = 0;

        int bit = 1;
        while (bit != 0) {
            if ((bit & intmask) != 0)
                len++;
            bit = bit << 1;
        }

        return len;
    }

    private static byte[] getAddress(String address) {//将字符串子网掩码转为byte数组
        String subStr[] = address.split("\\.");
        if (subStr.length != 4) throw new IllegalArgumentException("所传入的IP地址不符合IPv4的规范");
        byte b[] = new byte[4];
        for (int i = 0; i < b.length; i++) b[i] = (byte) Integer.parseInt(subStr[i]);
        return b;
    }
}
