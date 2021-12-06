// See https://aka.ms/new-console-template for more information

using SharpDocx;

var document = DocumentFactory.Create("Template.cs.docx");
document.Generate("output.docx");