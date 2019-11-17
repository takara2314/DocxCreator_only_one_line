using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace DocxCreator_only_one_line {
	class Program {
		static void Main(string[] args) {
			// インスタンス化して非staticメソッドを参照
			Program progra = new Program();
			progra.DocxCreate(@"テスト文章.docx");
		}

		void DocxCreate(string filepath) {
			Document document = new Document();
			Body body = new Body();
			Paragraph p = new Paragraph();
			Run run = new Run();


			RunProperties properties = new RunProperties();
			RunFonts fonts = new RunFonts() { Ascii = "ＭＳ Ｐゴシック", HighAnsi = "ＭＳ Ｐゴシック", EastAsia = "ＭＳ Ｐゴシック" };
			properties.Append(fonts);

			Text text = new Text() { Text = "こんにちは、世界！" };

			run.Append(properties);
			run.Append(text);


			p.Append(run);
			body.Append(p);
			document.Append(body);

			using (var package = WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document)) {
				MainDocumentPart mainDocumentPart = package.AddMainDocumentPart();
				mainDocumentPart.Document = document;
			}

			Console.WriteLine("docxファイルを生成しました。");
		}
	}
}
