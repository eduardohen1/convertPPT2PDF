package br.com.ehsolucoes.ConvertPPT2PDF;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImage;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PPTXToPDFConverter {

    private String pptxFilePath;
    private String pdfFilePath;

    public PPTXToPDFConverter(String pptxFilePath, String pdfFilePath) {
        this.pptxFilePath = pptxFilePath;
        this.pdfFilePath = pdfFilePath;
    }

    public String getPptxFilePath() {
        return pptxFilePath;
    }

    public void setPptxFilePath(String pptxFilePath) {
        this.pptxFilePath = pptxFilePath;
    }

    public String getPdfFilePath() {
        return pdfFilePath;
    }

    public void setPdfFilePath(String pdfFilePath) {
        this.pdfFilePath = pdfFilePath;
    }

    public Boolean convert() {
        if(pptxFilePath == null || pptxFilePath.isEmpty()) {
            System.out.println("Arquivo PPTX não informado!");
            return false;
        }
        if(pdfFilePath == null || pdfFilePath.isEmpty()) {
            System.out.println("Arquivo PDF não informado!");
            return false;
        }
        try {
            //caregar o arquivo
            FileInputStream inputStream = new FileInputStream(pptxFilePath);
            XMLSlideShow ppt = new XMLSlideShow(inputStream);
            inputStream.close();

            //Criar o arquivo de saida
            PDDocument pdf = new PDDocument();
            for(XSLFSlide slide : ppt.getSlides()){
                //definir a largura e altura da imagem
                int slideWidth = (int) ppt.getPageSize().getWidth();
                int slideHeight = (int) ppt.getPageSize().getHeight();

                //Criar uma nova página no PDF e definir seu tamanho
                PDPage page = new PDPage(new PDRectangle(slideWidth, slideHeight));
                pdf.addPage(page);
                PDPageContentStream content = new PDPageContentStream(pdf, page);

                //Converter o slide para uma imagem
                BufferedImage img = new BufferedImage(slideWidth, slideHeight, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = img.createGraphics();
                graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
                graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
                graphics.setColor(Color.white);
                graphics.clearRect(0, 0, slideWidth, slideHeight);
                slide.draw(graphics);

                //Adicionar a imagem ao PDF
                PDImageXObject pdImage = PDImageXObject.createFromByteArray(pdf, imageToByteArray(img), "slide");
                content.drawImage(pdImage, 0, 0);

                graphics.dispose();
                content.close();
            }

            //Salvar o PDF
            FileOutputStream outputStream = new FileOutputStream(pdfFilePath);
            pdf.save(outputStream);
            pdf.close();
            outputStream.close();

            System.out.println("Conversão concluída!");
            return true;
        }catch (Exception e) {
            System.out.println("Erro: " + e.getMessage());
            return false;
        }
    }

    private byte[] imageToByteArray(BufferedImage img) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(img, "png", baos);
        return baos.toByteArray();
    }

}
