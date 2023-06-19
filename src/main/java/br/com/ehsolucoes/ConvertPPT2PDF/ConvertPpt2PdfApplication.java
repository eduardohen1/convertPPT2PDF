package br.com.ehsolucoes.ConvertPPT2PDF;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ConvertPpt2PdfApplication {

	public static void main(String[] args) {
		SpringApplication.run(ConvertPpt2PdfApplication.class, args);

		PPTXToPDFConverter converter = new PPTXToPDFConverter("c:\\dev\\planoacao.pptx", "c:\\dev\\planoacao.pdf");
		converter.convert();
	}

}
