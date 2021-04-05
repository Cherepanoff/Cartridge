using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using Word=Microsoft.Office.Interop.Word;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Картриджи
{
    public partial class Zav : Form
    {
        private Word.Application wordapp;
        private Word.Documents worddocuments;
        private Word.Document worddocument;
        public Zav()
        {
            InitializeComponent();
        }

        private void Zav_Load(object sender, EventArgs e)
        {
            wordapp = new Word.Application();
            wordapp.Visible = true;
            Object template = Type.Missing;
            Object newTemplate = false;
            Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            Object visible = true;
            //Создаем документ 1
           wordapp.Documents.Add(
           ref template, ref newTemplate, ref documentType, ref visible);
            //Меняем шаблон
            template = @"C:\заявка картриджи от 10-09-2019.docx";
            //Создаем документ 2 worddocument в данном случае создаваемый объект 
            worddocument =
            wordapp.Documents.Add(
             ref template, ref newTemplate, ref documentType, ref visible);
        }
    }
}
