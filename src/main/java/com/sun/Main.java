package com.sun;

import com.aspose.words.Document;
import org.junit.Test;

/**
 * 测试
 * Created by sun'afei on 2019/03/22
 */
public class Main {

    @Test
    public void test1() throws Exception {
        Document document = new Document(this.getClass().getResourceAsStream("/沁园春-长沙.docx"));
        WatermarkUtils.insertWatermarkImg(document, "/defaultWatermark.png");
        document.save("c://out//图片水印.docx");
    }

    @Test
    public void test2() throws Exception {
        Document document = new Document(this.getClass().getResourceAsStream("/沁园春-雪.docx"));
        WatermarkUtils.insertWatermarkText(document, "我是文字水印");
        document.save("c://out//文字水印.docx");
    }
}
