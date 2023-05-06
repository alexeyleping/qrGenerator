package com.example.qrgenerator;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.WriterException;
import com.google.zxing.client.j2se.MatrixToImageConfig;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

@SpringBootApplication
public class QrGeneratorApplication implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(QrGeneratorApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        FileInputStream file = new FileInputStream("D:\\data.xlsx"); //path to Excel file(1 row name, 2 row data)
        XSSFWorkbook book = new XSSFWorkbook(file);
        XSSFSheet sheet = book.getSheet("Лист1"); //name sheet to book
        HashMap<String, String> map = new HashMap<>();
        for (int row = 0; row < sheet.getLastRowNum(); row++) {
            String key = sheet.getRow(row).getCell(0).getStringCellValue();
            String value = sheet.getRow(row).getCell(1).getStringCellValue();
            map.put(key, value);
        }

        Iterator<Map.Entry<String, String>> entryIterator = map.entrySet().iterator();
        while (entryIterator.hasNext()) {
            Map.Entry<String, String> mapEntry = entryIterator.next();
            qrGenerator(mapEntry.getKey(), mapEntry.getValue());
        }
        book.close();
        file.close();
    }

    private void qrGenerator(String nameQrFile, String dataValue) throws WriterException, IOException {
        QRCodeWriter qrCodeWriter = new QRCodeWriter();
        BitMatrix bitMatrix = qrCodeWriter.encode(dataValue, BarcodeFormat.QR_CODE, 300, 300);
        MatrixToImageConfig matrixToImageConfig = new MatrixToImageConfig(-1, 0xFF000000);
        BufferedImage bufferedImage = MatrixToImageWriter.toBufferedImage(bitMatrix, matrixToImageConfig);
        BufferedImage logo = ImageIO.read(new File("D:\\logo.xlsx")); //logo file(size 40*40 or else)
        int w = Math.max(bitMatrix.getWidth(), logo.getWidth());
        int h = Math.max(bitMatrix.getHeight(), logo.getHeight());
        BufferedImage combined = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
        Graphics graphics = combined.getGraphics();
        graphics.drawImage(bufferedImage, 0, 0, null);
        graphics.drawImage(logo, 130, 130, null);
        ImageIO.write(combined, "PNG", new File("D:\\qr\\" + nameQrFile + ".png"));
    }
}
