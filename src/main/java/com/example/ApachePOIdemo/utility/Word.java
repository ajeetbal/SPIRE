package com.example.ApachePOIdemo.utility;

import java.io.IOException;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.TextSelection;

public class Word {

	public static void main(String[] args) throws IOException {
		String input = "/home/ajeet/Music/demo.docx";
		String output = "/home/ajeet/Music/replaceWithText.docx";
		Document document = new Document();
		document.loadFromFile(input, FileFormat.Docx);

		document.replace("Word", "NewWord", false, true);
		TextSelection[] organizationName = document.findAllString("<Organization Name>", false, true);
		TextSelection[] CMMC = document.findAllString("“${cmmc-version}”", false, true);
		TextSelection[] DOCUMENT = document.findAllString("“${document-version}”", false, true);
		TextSelection[] Organ = document.findAllString("“${organization-name}”", false, true);
		for (TextSelection selection : Organ) {
			selection.getAsOneRange().setText("Reliance");
		}

		for (TextSelection selection : organizationName) {
			selection.getAsOneRange().setText("Reliance");
		}
		for (TextSelection selection : CMMC) {
			selection.getAsOneRange().setText("CMMC_V3");
		}
		for (TextSelection selection : DOCUMENT) {
			selection.getAsOneRange().setText("DOCUMENT_C2");
		}

		document.saveToFile(output, FileFormat.Docx);
	}

}