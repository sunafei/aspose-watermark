package com.sun;

import com.aspose.words.*;
import com.aspose.words.Shape;
import org.apache.commons.lang3.StringUtils;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Random;

/**
 * 水印工具类
 * Created by sun'afei on 2019/03/22
 */
public class WatermarkUtils {

    /**
     * 水印每个字体的宽度为10
     */
    private final static int WATERMARK_FONT_WIDTH = 10;
    /**
     * 水印每个字体的高度为16
     */
    private final static int WATERMARK_FONT_HEIGHT = 16;
    /**
     * 水印图片的默认高度为20
     */
    private final static double WATERMARK_IMG_HEIGHT = 20d;
    /**
     * 水印倾斜角度 默认是50 为保证文字连续性尽量不要修改
     */
    private final static int WATERMARK_FONT_ROTATION = -50;
    /**
     * 水印字体颜色
     */
    private static Color watermark_color = new Color(217, 226, 243);

    /**
     * 注册apose.words
     */
    static {
        try {
            InputStream is = new FileInputStream("");
            License aposeLic = new License();
            aposeLic.setLicense(is);
        } catch (Exception e) {
            System.out.println("apose注册失败,会导致系统打印呈现未注册状态,文档上方出现未注册红字!");
        }
    }

    /**
     * word文档插入图片水印
     *
     * @param doc     文档对象
     * @param imgPath 图片路径
     */
    public static void insertWatermarkImg(Document doc, String imgPath) {
        if (StringUtils.isBlank(imgPath)) {
            System.out.println("没有配置水印图片, 无法为文档加入水印");
            return;
        }
        Paragraph watermarkPara = new Paragraph(doc);
        // TODO 这里的数据 计算水印个数(900 150 700 150) 首个水印位置(-200至-100)都是实验得到 没有理论依据
        for (int top = 0; top < 900; top += 150) {
            int beginLeft = new Random().ints(-100, -50).limit(1).findFirst().getAsInt();
            for (int left = beginLeft; left < 700; left += 150) {
                Shape waterShape = buildImgShape(doc, imgPath, left, top);
                watermarkPara.appendChild(waterShape);
            }
        }
        // 在每个部分中，最多可以有三个不同的标题，因为我们想要出现在所有页面上的水印，插入到所有标题中。
        for (Section sect : doc.getSections()) {
            // 每个区段可能有多达三个不同的标题，因为我们希望所有页面上都有水印，将所有的头插入。
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }
    }

    /**
     * 构建图片shape类
     *
     * @param doc     文档对象
     * @param imgPath 图片文件路径
     * @param left    左边距
     * @param top     上边距
     */
    private static Shape buildImgShape(Document doc, String imgPath, int left, int top) {
        Shape shape = new Shape(doc, ShapeType.IMAGE);
        try {
//            File imgFile = new File(imgPath);
//            BufferedImage sourceImg = ImageIO.read(new FileInputStream(imgFile));
            BufferedImage sourceImg = ImageIO.read(new WatermarkUtils().getClass().getResourceAsStream(imgPath));
            double multiple = sourceImg.getHeight() / WATERMARK_IMG_HEIGHT;
//            shape.setWidth(sourceImg.getWidth() / multiple);
//            System.out.println(sourceImg.getWidth() / multiple);
//            shape.setHeight(WATERMARK_IMG_HEIGHT);
            shape.getImageData().setImage(sourceImg);
            shape.setWidth(50);
            shape.setHeight(20);
            shape.setRotation(WATERMARK_FONT_ROTATION);
            shape.setLeft(left);
            shape.setTop(top);
            shape.setWrapType(WrapType.NONE);
        } catch (Exception e) {
            throw new RuntimeException("图片附件丢失, 无法生成水印!", e);
        }
        return shape;
    }

    /**
     * 为word文档插入文本水印
     *
     * @param doc           文档对象
     * @param watermarkText 文本内容
     */
    public static void insertWatermarkText(Document doc, String watermarkText) {
        if (StringUtils.isBlank(watermarkText)) {
            System.out.println("没有配置水印内容, 无法为文档加入水印");
            return;
        }
        Paragraph watermarkPara = new Paragraph(doc);
        // TODO 这里的数据 计算水印个数(900 150 700 150) 首个水印位置(-200至-100)都是实验得到 没有理论依据
        for (int top = 0; top < 900; top += 150) {
            int beginLeft = new Random().ints(-200, -100).limit(1).findFirst().getAsInt();
            for (int left = beginLeft; left < 700; left += 150) {
                // 如果是左起第一个水印则通过字符串截取达到随机显示水印的谜底
                // 这样做的原因为了保证倾斜的行保证对齐 又能表现随机的特性 不是好办法
                if (left == beginLeft) {
                    int splitNo = new Random().ints(0, watermarkText.length()).limit(1).findFirst().getAsInt();
                    String firstWater = watermarkText.substring(splitNo) + "                                            ".substring(0, splitNo);
                    Shape waterShape = buildTextShape(doc, firstWater, left, top);
                    watermarkPara.appendChild(waterShape);
                } else {
                    Shape waterShape = buildTextShape(doc, watermarkText, left, top);
                    watermarkPara.appendChild(waterShape);
                }
            }
        }

        // 在每个部分中，最多可以有三个不同的标题，因为我们想要出现在所有页面上的水印，插入到所有标题中。
        for (Section sect : doc.getSections()) {
            // 每个区段可能有多达三个不同的标题，因为我们希望所有页面上都有水印，将所有的头插入。
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_PRIMARY);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_FIRST);
            insertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HEADER_EVEN);
        }

    }

    /**
     * 构建文字shape类
     *
     * @param doc           文档对象
     * @param watermarkText 水印文字
     * @param left          左边距
     * @param top           上边距
     */
    private static Shape buildTextShape(Document doc, String watermarkText, double left, double top) {
        Shape watermark = new Shape(doc, ShapeType.TEXT_PLAIN_TEXT);
        watermark.getTextPath().setText(watermarkText);
        watermark.getTextPath().setFontFamily("宋体");
        watermark.setWidth(watermarkText.length() * WATERMARK_FONT_WIDTH);
        watermark.setHeight(WATERMARK_FONT_HEIGHT);
        watermark.setRotation(WATERMARK_FONT_ROTATION);
        //绘制水印颜色
        watermark.getFill().setColor(watermark_color);
        watermark.setStrokeColor(watermark_color);
        //将水印放置在页面中心
        watermark.setLeft(left);
        watermark.setTop(top);
        watermark.setWrapType(WrapType.NONE);
        return watermark;
    }

    /**
     * 插入水印
     *
     * @param watermarkPara 水印段落
     * @param sect          部件
     * @param headerType    头标类型字段
     */
    private static void insertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, int headerType) {
        HeaderFooter header = sect.getHeadersFooters().getByHeaderFooterType(headerType);
        if (header == null) {
            header = new HeaderFooter(sect.getDocument(), headerType);
            sect.getHeadersFooters().add(header);
        }
        header.appendChild(watermarkPara.deepClone(true));
    }
}
