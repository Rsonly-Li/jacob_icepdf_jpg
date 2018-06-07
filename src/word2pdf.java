import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import org.icepdf.core.pobjects.Document;
import org.icepdf.core.pobjects.Page;
import org.icepdf.core.util.GraphicsRenderingHints;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;


/**
 * Created by 李綦睿 on 2017-03-13 0013.
 */
public class word2pdf {
    static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
    static final int wdFormatPDF = 17;// word转PDF 格式

    public static void main(String[] args) throws Exception {
        String source0 = "E:\\temp\\试题导入.docx";
        String target1 = "E:\\temp\\test0.pdf";
        String source1 = "111.png";
        word2pdf pdf = new word2pdf();
        pdf.word2pdf(source0, target1);
        //pdf.icepdf2png(target1);
        pdf.icepdf2jpg(target1);



        //pdf.PictureProcess();
        //File fromFile = new File("E:\\temp\\imageCapture1_4.png");
        //File toFile = new File("E:\\temp\\imageCapture1_444444.png");
        //resizePng(fromFile, toFile, 100, 100, false);
        //List<String> list1 = new ArrayList<String>();
        //list1.add("1");
        //list1.add("2");
        //list1.add("3");
        //list1.add("4");
        //List<String> list2 = new ArrayList<String>();
        //list2.add("a");
        //list2.add("b");
        //list2.add("c");
        //list2.add("d");
        //list1.addAll(list2);
        //list1.addAll(list1);
        //System.out.println(list1);
        //List<Integer> number = new ArrayList<Integer>();
        //if (null==source1 || source1.equals("")){
        //    System.out.println("错误！");
        //}else {
        //    for (int i = 0;i < source1.length();i++){
        //        if (source1.charAt(i)=='.'){
        //            System.out.println(i);
        //            number.add(i);
        //        }
        //    }
        //    System.out.println(number);
        //    String type = source1.substring(number.get(number.size()-1));
        //    System.out.println(type);
        //}
        /*List<String> name = new ArrayList<String>();
        name.add("李綦睿");
        name.add("哈哈");
        name.add("嘿嘿");
        name.add("哼哼");
        name.add("李綦睿(1)");
        name.add("李綦睿(2)");
        name.add("李綦睿(4)");
        name.add("(李綦睿)");
        name.add("李(綦)睿");
        name.add("李(綦睿)");
        String s = "李(綦睿)";
        System.err.println(s.indexOf("("));
        String ss = autoRename(s, name, 0);
        System.err.println(ss);*/

        /*String oldName = "(李綦睿)";
        if (CollectionUtils.isNotEmpty(existList)) {
            System.out.println(getName(oldName, existList, 1));
        } else {
            System.out.println(oldName);
        }*/

        //InputStream inputStream = getClass().getResourceAsStream("libra.properties");

        //Properties properties = new MyProperties();
        //String filepath = "D:\\IDEA workpaces\\jacob_icepdf_jpg\\src\\libra.properties";
        //File file = new File(filepath);
        //if (!file.exists())
        //    file.createNewFile();
        //FileInputStream fis = new FileInputStream(file);
        //properties.load(fis);
        //System.out.println(properties.get("jdbc.oracle.password"));
        //Iterator<String> iterator = properties.stringPropertyNames().iterator();
        //while (iterator.hasNext()){
        //    String key = iterator.next();
        //    System.out.println(key +":"+ properties.getProperty(key));
        //}
        //fis.close();
        //
        //FileOutputStream fos = new FileOutputStream(file);
        //properties.setProperty("jdbc.oracle.password", "10086");
        //properties.store(fos, null);
        //fos.close();
        //
        //MyProperties myProperties = new MyProperties();
        //String url1 = myProperties.getUrl();
        //System.out.println(url1);


        //FileInputStream fis11 = new FileInputStream(file);
        //properties.load(fis11);
        //System.out.println(properties.get("jdbc.oracle.password"));
        //fis11.close();
        //FileOutputStream fos11 = new FileOutputStream(file);
        //properties.store(fos11, null);
        //fos11.close();

    }

    private static List<String> existList;

    static {
        String name = "李綦睿";
        String name2 = "李綦睿(1)";
        String name3 = "李綦睿(2)";
        String name6 = "李綦睿(4)";
        String name4 = "李綦睿(1)(1)";
        String name5 = "李綦睿(1)(2)";
        String name7 = "(李綦睿)";
        existList = new ArrayList<String>();
        existList.add(name);
        existList.add(name2);
        existList.add(name3);
        existList.add(name4);
        existList.add(name5);
        existList.add(name6);
        existList.add(name7);
    }


    private static String getName(String oldName, List<String> existNameList, int index) {
        String result = null;
        if (existNameList.contains(oldName)) {
            result = oldName + "(" + index + ")";
            if (existNameList.contains(result)) {
                result = getName(oldName, existNameList, ++index);
            }
        } else {
            result = oldName;
        }
        return result;
    }



    public static String autoRename(String s, List<String> name, int x) {
        String s1 = null;
        if (name.contains(s)) {
            if (s.indexOf("(") != -1) {
                if (name.contains(s)) {
                    s1 = s.substring(0, s.indexOf("(")) + "(" + (x + 1) + ")";
                    x++;
                    return autoRename(s1, name, x);
                } else {
                    s1 = s;
                    return s1;
                }
            } else {
                if (name.contains(s)) {
                    s1 = s + "(" + (x + 1) + ")";
                    x++;
                    return autoRename(s1, name, x);
                } else {
                    s1 = s;
                    return s1;
                }
            }
        } else {
            s1 = s;
        }
        return s1;
    }


    /**
     * jacob将word转pdf
     *
     * @param source
     * @param target
     * @return
     */
    public static boolean word2pdf(String source, String target) {
        System.out.println("Word转PDF开始启动...");
        long start = System.currentTimeMillis();//word转换pdf开始时间
        ActiveXComponent app = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", false);
            Dispatch docs = app.getProperty("Documents").toDispatch();
            System.out.println("打开文档：" + source);
            Dispatch doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();
            System.out.println("转换文档到PDF：" + target);
            File tofile = new File(target);
            if (tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(doc, "SaveAs", target, wdFormatPDF);
            Dispatch.call(doc, "Close", false);
            long end = System.currentTimeMillis();//转换结束时间
            System.out.println("转换完成，用时：" + (end - start) + "ms");
            return true;
        } catch (Exception e) {
            System.out.println("Word转PDF出错：" + e.getMessage());
            return false;
        } finally {
            if (app != null) {
                app.invoke("Quit", wdDoNotSaveChanges);
            }
        }
    }

    public static void icepdf2png(String path) {
        String filePath = path;//pdf文件路径
        Document document = new Document();//pdf文件对象
        word2pdf pdf = new word2pdf();
        try {
            document.setFile(filePath);

            float scale = 1.5f;//缩放比例

            float rotation = 0f;//旋转角度

            for (int i = 0; i < document.getNumberOfPages(); i++) {
                BufferedImage image = (BufferedImage) document.getPageImage(i, GraphicsRenderingHints.SCREEN, Page.BOUNDARY_CROPBOX, rotation, scale);
                RenderedImage renderedImage = image;
                try {
                    System.out.println(" capturing page " + i);
                    File file = new File("E:\\temp\\imageCapture1_" + i + ".png");
                    String imagePath = "E:\\temp\\imageCapture1_" + i + ".png";
                    String imageSavePath = "E:\\temp\\imageCapture1_new" + i + ".png";
                    ImageIO.write(renderedImage, "png", file);
                    pdf.PictureProcess(imagePath, imageSavePath);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                image.flush();
            }
            document.dispose();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void icepdf2jpg(String path){
        String filePath = path;//pdf文件路径
        Document document = new Document();
        word2pdf pdf = new word2pdf();
        try {
            document.setFile(filePath);
            float scale = 1.5f;
            float rotation = 0f;
            for (int i = 0;i < document.getNumberOfPages();i++){
                BufferedImage image = (BufferedImage)document.getPageImage(i,GraphicsRenderingHints.SCREEN,Page.BOUNDARY_CROPBOX,rotation,scale);
                RenderedImage renderedImage = image;
                try {
                    System.out.println(" capturing page " + i);
                    File file = new File("E:\\temp\\imageCapture1_" + i + ".jpg");
                    ImageIO.write(renderedImage,"jpg",file);
                }catch (Exception e){
                    e.printStackTrace();
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }


    /**
     * 像素扫描
     *
     * @param list
     * @return
     */
    private static boolean same(List<Integer> list) {
        System.err.println(list);
        if (list == null || list.size() == 0) {
            System.err.println(true);
            return true;
        }
        return new HashSet<Integer>(list).size() <= 3;
    }
    /*private static boolean same(List<Integer> list) {
        if (list == null || list.size() == 0) {
            return true;
        }
        Integer first = list.get(0);
        for (int i = 1; i < list.size(); i++) {
            if (Math.abs(list.get(i) - first) > 10) {
                return false;
            }
        }
        return true;
    }*/


    /**
     * 带参方法，集成使用
     *
     * @param s
     * @param ss
     * @throws Exception
     */
    public void PictureProcess(String s, String ss) throws Exception {
        BufferedImage image = ImageIO.read(new File(s));

        int startRows = 0;
        int endRows = image.getHeight();
        int startCol = 0;
        int endCol = image.getWidth();

        while (true) {
            int _startRows = startRows;//上边界
            int _endRows = endRows;//下边界
            int _startCol = startCol;//左边界
            int _endCol = endCol;//右边界
            // 从上到下
            for (int i = startRows; i < endRows; i++) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startCol; j < endCol; j++) {
                    list.add(image.getRGB(j, i));
                }
                if (same(list)) {
                    startRows = i;
                } else {
                    break;
                }
            }
            // 从下到上
            for (int i = endRows - 1; i >= startRows; i--) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startCol; j < endCol; j++) {
                    list.add(image.getRGB(j, i));
                }
                if (same(list)) {
                    endRows = i;
                } else {
                    break;
                }
            }

            // 从左到右
            for (int i = startCol; i < endCol; i++) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startRows; j < endRows; j++) {
                    list.add(image.getRGB(i, j));
                }
                if (same(list)) {
                    startCol = i;
                } else {
                    break;
                }
            }
            // 从右到左
            for (int i = endCol - 1; i >= startCol; i--) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startRows; j < endRows; j++) {
                    list.add(image.getRGB(i, j));
                }
                if (same(list)) {
                    endCol = i;
                } else {
                    break;
                }
            }

            if (_startRows == startRows && _endRows == endRows
                    && _startCol == startCol && _endCol == endCol) {
                break;
            }
        }

        BufferedImage imageOut = new BufferedImage(endCol - startCol,
                endRows - startRows, image.getType());
        for (int i = 0; i < imageOut.getWidth(); i++) {
            for (int j = 0; j < imageOut.getHeight(); j++) {
                imageOut.setRGB(i, j,
                        image.getRGB(i + startCol, j + startRows));
            }
        }
        ImageIO.write(imageOut, "png", new File(ss));
        System.err.println("图片裁剪完成！");

    }


    /**
     * 不带参数的方法，单纯调试剪切功能使用
     *
     * @throws Exception
     */
    public void PictureProcess() throws Exception {
        BufferedImage image = ImageIO.read(new File("E:\\temp\\imageCapture1_4.jpg"));

        int startRows = 0;
        int endRows = image.getHeight();
        int startCol = 0;
        int endCol = image.getWidth();

        while (true) {
            int _startRows = startRows;
            int _endRows = endRows;
            int _startCol = startCol;
            int _endCol = endCol;
            // 从上到下
            for (int i = startRows; i < endRows; i++) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startCol; j < endCol; j++) {
                    list.add(image.getRGB(j, i));
                }
                if (same(list)) {
                    startRows = i;
                } else {
                    break;
                }
            }
            // 从下到上
            for (int i = endRows - 1; i >= startRows; i--) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startCol; j < endCol; j++) {
                    list.add(image.getRGB(j, i));
                }
                if (same(list)) {
                    endRows = i;
                } else {
                    break;
                }
            }

            // 从左到右
            for (int i = startCol; i < endCol; i++) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startRows; j < endRows; j++) {
                    list.add(image.getRGB(i, j));
                }
                if (same(list)) {
                    startCol = i;
                } else {
                    break;
                }
            }
            // 从右到左
            for (int i = endCol - 1; i >= startCol; i--) {
                List<Integer> list = new ArrayList<Integer>();
                for (int j = startRows; j < endRows; j++) {
                    list.add(image.getRGB(i, j));
                }
                if (same(list)) {
                    endCol = i;
                } else {
                    break;
                }
            }

            if (_startRows == startRows && _endRows == endRows
                    && _startCol == startCol && _endCol == endCol) {
                break;
            }
        }

        BufferedImage imageOut = new BufferedImage(endCol - startCol,
                endRows - startRows, image.getType());
        for (int i = 0; i < imageOut.getWidth(); i++) {
            for (int j = 0; j < imageOut.getHeight(); j++) {
                imageOut.setRGB(i, j,
                        image.getRGB(i + startCol, j + startRows));
            }
        }
        ImageIO.write(imageOut, "jpg", new File("E:\\temp\\out0.jpg"));

    }

    /**
     * 裁剪PNG图片工具类
     *
     * @param fromFile
     *            源文件
     * @param toFile
     *            裁剪后的文件
     * @param outputWidth
     *            裁剪宽度
     * @param outputHeight
     *            裁剪高度
     * @param proportion
     *            是否是等比缩放
     */
    //public static void resizePng(File fromFile, File toFile, int outputWidth, int outputHeight,
    //                             boolean proportion) {
    //    try {
    //        BufferedImage bi2 = ImageIO.read(fromFile);
    //        int newWidth;
    //        int newHeight;
    //        // 判断是否是等比缩放
    //        if (proportion) {
    //            // 为等比缩放计算输出的图片宽度及高度
    //            double rate1 = ((double) bi2.getWidth(null)) / (double) outputWidth + 0.1;
    //            double rate2 = ((double) bi2.getHeight(null)) / (double) outputHeight + 0.1;
    //            // 根据缩放比率大的进行缩放控制
    //            double rate = rate1 < rate2 ? rate1 : rate2;
    //            newWidth = (int) (((double) bi2.getWidth(null)) / rate);
    //            newHeight = (int) (((double) bi2.getHeight(null)) / rate);
    //        } else {
    //            newWidth = outputWidth; // 输出的图片宽度
    //            newHeight = outputHeight; // 输出的图片高度
    //        }
    //        BufferedImage to = new BufferedImage(newWidth, newHeight, BufferedImage.TYPE_INT_RGB);
    //        Graphics2D g2d = to.createGraphics();
    //        to = g2d.getDeviceConfiguration().createCompatibleImage(newWidth, newHeight,
    //                Transparency.TRANSLUCENT);
    //        g2d.dispose();
    //        g2d = to.createGraphics();
    //        @SuppressWarnings("static-access")
    //        Image from = bi2.getScaledInstance(newWidth, newHeight, bi2.SCALE_AREA_AVERAGING);
    //        g2d.drawImage(from, 0, 0, null);
    //        g2d.dispose();
    //        ImageIO.write(to, "png", toFile);
    //    } catch (Exception e) {
    //        e.printStackTrace();
    //    }
    //}
}
