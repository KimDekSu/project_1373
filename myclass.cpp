#include "myclass.h"
#include "ui_myclass.h"

myclass::myclass(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::myclass)
{
    ui->setupUi(this);
}

myclass::~myclass()
{
    delete ui;
}

bool myclass::setPath(QString path){
    if(!path.isEmpty()){
        this->path = path;
        if(OpenXML(this->path)==false){
            return false;
        }
    }
    return true;
}


void myclass::on_actionOpen_File_triggered()// открытие файла из панели инструментов
{
    // QFileDialog::getOpenFileName - получает путь к файлу (параметры copypaste из документации)
    path = QFileDialog::getOpenFileName(this, "Open File", "C:/",
                                        "All Files(*.*);; DOCX (*.docx);; ODT (*.odt)");
    if (!path.isEmpty()){
        OpenXML(path);
    }
}





bool myclass::OpenXML(QString path)
{
    ui->textEdit->clear(); // очистка окна перед открытием нового документа

    QString file_exten;
    if (path.right(4) == "docx") { file_exten = "docx"; }
    else if(path.right(3) == "odt"){ file_exten = "odt"; }
    //вывод окна с ошибкой при попытке открыть файл неправильного формата (не docx и не odt)
    else QMessageBox::critical(nullptr,"Error::Invalid file format","The application can only\nopen .docx and .odt files");

    QByteArray dataContent, dataStyles, dataImageName, dataNumbering; // xml-документы для парсинга
    QZipReader zip(path, QIODevice::ReadOnly);
    if (zip.status() == QZipReader::NoError)
    {
        if (file_exten == "docx"){
            // get .docx-file
            dataContent = zip.fileData("word/document.xml"); // основное содержание docx-файла
            dataStyles = zip.fileData("word/styles.xml"); // стили, используемые в docx-файле
            dataNumbering = zip.fileData("word/numbering.xml"); // оформление списков, используемых в docx-файле
            dataImageName = zip.fileData("word/_rels/document.xml.rels"); // информация о картинках в docx-файле
        }else {
            // get .odt-file
            dataContent = zip.fileData("content.xml");// основное содержание odt-файла
            dataStyles = zip.fileData("styles.xml");// стили, используемые в odt-файле
        }
        zip.close();
        // выгружаем содержимое xml-файла в объект класса QDomDocument
        document.setContent(dataContent);
        stylesInfo.setContent(dataStyles);
        listInfo.setContent(dataNumbering);
        imageName.setContent(dataImageName);

        ParsXML(document, file_exten); // запуск парсинга документа
        return true;
    }
    else
    {
        qDebug() << "Faild zip";
        // вывод окна с ошибкой при неправильно заданном пути к файлу
        QMessageBox::critical(nullptr,"Error::File not found","Check the file path.\nFile does not exist.");
        return false;
    }
}

void myclass::ParsXML(QDomDocument document, QString file_exten)
{
    QDomElement root = document.firstChildElement(); // получение корня (xml представим в виде дерева)
    QDomElement node = root.firstChild().toElement(); // переход к первому ребенку

    if  (file_exten == "docx"){
        // заполнение списка используемых тегов
        TextTagsNameQList.push_back("w:t"); // текст
        TextTagsNameQList.push_back("w:rFonts"); // шрифт
        TextTagsNameQList.push_back("w:b"); // жирное выделение
        TextTagsNameQList.push_back("w:i"); // кусив
        TextTagsNameQList.push_back("w:u"); // почеркивание
        TextTagsNameQList.push_back("w:pStyle"); // стиль форматирования текста (может содержать информацию нескольких тегов, хранится в отдельном документе
        TextTagsNameQList.push_back("w:sz"); // размер шрифта
        TextTagsNameQList.push_back("w:color"); // цвет текста
        TextTagsNameQList.push_back("w:highlight"); // цветное выделение текста (текстовыделитель)
        TextTagsNameQList.push_back("w:spacing"); // межстрочные интервалы
        TextTagsNameQList.push_back("w:jc"); // выравнивания текста
        TextTagsNameQList.push_back("w:ind"); // отступы
        TextTagsNameQList.push_back("w:p"); // тег, обозначающий абзац
        TextTagsNameQList.push_back("a:blip"); // id изображения
        // for numbering
        TextTagsNameQList.push_back("w:ilvl"); // уровень списка
        TextTagsNameQList.push_back("w:numFmt"); // стиль стписка (нумерованный, маркерованный)
        TextTagsNameQList.push_back("w:lvlJc"); // выравнивания текста

        paragNode var_node; // переменная для хранения форматирования текста
        ClearNode(&var_node); // инициализация
        GetDataFromDOCX(node, &var_node); // начало работы с обходом и тегами
        TextTagsNameQList.clear(); // очистка списка тегов
    }else{
        TextTagsNameQList.push_back("text:p");//тег абзаца
        TextTagsNameQList.push_back("text:h");//тег заголовка
        TextTagsNameQList.push_back("text:span");//тег текста
        TextTagsNameQList.push_back("draw:image");//тег изображения
        TextTagsNameQList.push_back("style:text-properties");//тег стилей текста
        TextTagsNameQList.push_back("text:s");//тег пробела
        TextTagsNameQList.push_back("style:paragraph-properties");//тег стилей абзаца
        TextTagsNameQList.push_back("text:list-style");//тег стилей списке
        TextTagsNameQList.push_back("text:list-item");

        paragNode var_node;// переменная для хранения форматирования текста
        ClearNode(&var_node); // инициализация
        ClearNode(&old_var_node);// инициализация переменной для хранения стилей абзаца
        GetDataFromODT(node, &var_node);// начало работы с обходом и тегами
        TextTagsNameQList.clear();// очистка списка тегов
    }
}

void myclass::GetDataFromDOCX(QDomElement node, paragNode *var_node) /* рекурсивная функция для обхода дерева и вывода отформатированного
текста. Реализован обход на основе прямого. Сначала идет функция идет вниз по дереву по child-элементам, пока не окажется в листе (первая
рекурсия закончилась). После этого возвращается в пройденные узлы и, в зависимости от тега узла, сохраняет заполненную структуру paragNode (с
текстом и форматированием) либо выводит на экран отформатированный текст. После чего осуществляется второй рекурсивный переход к
соседнему (по отношению к текущему) узлу.*/
{
    if (!node.toElement().isNull()){
        if (node.toElement().tagName() == "w:tbl"){     // обход дерева таблицы
            if (!node.nextSiblingElement().isNull()){
                int rows = GetRowsDOCX(node);
                int columns = GetColumnsDOCX(node);
                QTextCursor cursor(ui->textEdit->textCursor());
                QTextTableFormat format;
                format.setBorderCollapse(true);
                format.setCellPadding(15);
                format.setTopMargin(var_node->indentParag.top);
                format.setBottomMargin(var_node->indentParag.bottom);
                format.setLeftMargin(var_node->indentParag.left);
                format.setRightMargin(var_node->indentParag.right);
                QTextTable *table = cursor.insertTable(rows, columns, format);
                ParsTableDOCX(node, table, rows, columns);      //заполнение таблицы текстом и его форматирование

                GetDataFromDOCX(node.nextSiblingElement(), var_node);
            }
            return;
        }

        SearchTagDOCX(node, var_node); // выбор тега и заполнение структуры paragNode в зависимости от него. Функция switch-case

        if (!node.firstChildElement().isNull()){
            GetDataFromDOCX(node.firstChildElement(), var_node); // первый рекурсивный переход к child-элементу
        }

        if (node.tagName() == "w:pPr"){ // запоминание форматирования для всего абзаца
            old_var_node = *var_node;
        }

        if (node.tagName() == "w:r"){ // запоминание форматирования для отдельной части абзаца и восстановление форматирования для всего абзаца
            paragComponentsList.push_back(*var_node);
            *var_node = old_var_node;
        }
        if (node.tagName() == "w:p" && var_node->text.isEmpty()){ //  случай, когда в абзаце не оказалось текста
            var_node->text = " ";
            paragComponentsList.push_back(*var_node);
            ClearNode(var_node);
            ClearNode(&old_var_node);
        }

        if (node.tagName() == "w:p" && !paragComponentsList.isEmpty()){ // случай, когда в абзаце есть текст
            while (!paragComponentsList.isEmpty()){
                *var_node = paragComponentsList.front();
                paragComponentsList.pop_front();

                QTextCursor cursor = QTextCursor(ui->textEdit->textCursor());
                QTextBlockFormat blockFormat = cursor.blockFormat();

                // установака отступов, межстрочных интервалов, отступа абзаца и выравнивания для абзаца
                if (var_node->newParag){
                    blockFormat.setLeftMargin(var_node->indentParag.left);
                    blockFormat.setRightMargin(var_node->indentParag.right);
                    blockFormat.setTopMargin(var_node->indentParag.top);
                    blockFormat.setBottomMargin(var_node->indentParag.bottom);
                    blockFormat.setTextIndent(var_node->indentParag.firstLine);
                    blockFormat.setAlignment(var_node->alignParag);
                }
                cursor.setBlockFormat(blockFormat);

                // установка форматирования текста
                QTextCharFormat charFormat;
                charFormat.setFont(var_node->font);
                charFormat.setForeground(QColor(var_node->colorText));
                charFormat.setBackground(QColor(var_node->colorBackground));
                cursor.insertText(var_node->text, charFormat);
                if (!var_node->image_id.isEmpty()){ // вывод изображения
                    PrintImageDocx(var_node);
                }
            }
            ui->textEdit->append("");
            ClearNode(var_node);
        }
        if (!node.nextSiblingElement().isNull()){
             GetDataFromDOCX(node.nextSiblingElement(), var_node); // второй рекурсивный переход к соседнему (по отношению к текущему) узлу
        }
    }
}

void myclass::ParsTableDOCX(QDomElement node,QTextTable *table, int rows, int columns) /*нерекурсивная функция для обхода дерева таблицы
 для формата docx, заполению таблицы по ячейкам, форматированию текста в этих ячейках. Сначала ищется 1 строка и соотвественоо ее 1 элемент,
затем углубляясь внутрь табличного дерева находится поле с текстом и его атрибутами. Вызывается вспомогательный метод обработки атрибутов,
затем все вносится формат и вставляет ячейка с уже заполненным форматом*/
{
    QDomElement row = node.firstChild().toElement();
    QDomElement cell;
    QDomElement cellText;
    QString cellStr;
    QTextTableCell cellValue;
    QTextCursor cellCursor;
    paragNode var_node;

    while(row.tagName() != "w:tr")
    {
        row = row.nextSiblingElement();
    }
    cell = row.firstChild().toElement();

    for(int i = 0; i < rows; ++i) // w:tr
    {

        for(int j = 0; j < columns; ++j) // w:tc
        {
             if(cell.nextSiblingElement().isNull()) // переход к следующей строке
             {
                row = row.nextSiblingElement();
                cell = row.firstChild().toElement();
             }

             if(cell.tagName() != "w:tc") // переход к следующей ячейке
             {
                while(cell.tagName() != "w:tc")
                {
                    cell = cell.nextSiblingElement();
                }
             }
             else
             {
                if(j != 0)
                {
                    cell = cell.nextSiblingElement();
                }
             }

             cellText = cell.firstChild().toElement();
             while(cellText.tagName() != "w:p")
             {
                cellText = cellText.nextSiblingElement();
             }

             cellText = cellText.firstChild().toElement();
             while(cellText.tagName() != "w:r")
             {
                cellText = cellText.nextSiblingElement();
                if(cellText.isNull())
                {
                    break;
                }
             }
             // обработка шрифтов
             QDomElement fonts = cellText.firstChildElement().firstChildElement();
             ClearNode(&var_node);
             QTextCharFormat charFormat;
             if(cellText.tagName() == "w:r")
             {
                while(!fonts.isNull())
                {
                    SearchTagDOCX(fonts, &var_node);
                    fonts = fonts.nextSiblingElement();
                }
             }
             charFormat.setFont(var_node.font);
             charFormat.setForeground(QColor(var_node.colorText));
             charFormat.setBackground(QColor(var_node.colorBackground));

             cellText = cellText.firstChild().toElement();
             while(cellText.tagName() != "w:t" && !cellText.isNull())
             {
                cellText = cellText.nextSiblingElement();
                if(cellText.isNull())
                {
                    break;
                }
             }

             cellValue = table->cellAt(i,j);
             cellCursor = cellValue.firstCursorPosition();

             if (cellText.isNull()){
                cellCursor.insertText(" ");
             }else {
                cellCursor.insertText(cellText.text(), charFormat);
             }

        }
    }

}

int myclass::GetRowsDOCX(QDomElement node) /*вспомогательная функция для поиска количества строк в таблице для формата docx*/
{
    QDomElement child = node.firstChildElement();
    int rows = 0;
    while(!child.isNull())
    {
        if(child.tagName() == "w:tr")
        {
             ++rows;
        }
        child = child.nextSiblingElement();
    }
    return rows;
}

int myclass::GetColumnsDOCX(QDomElement node) /*вспомогательная функция для поиска количества столбцов в таблице для формата docx*/
{
    QDomNode child = node.firstChild();
    QDomNodeList children;
    int columns = 0;
    while(!child.isNull())
    {
        if(child.toElement().tagName() == "w:tblGrid")
        {
             children = child.childNodes();
             columns = children.length();
             return columns;
        }
        child = child.nextSibling();
    }
    return 0;
}

void myclass::ParsTableODT(QDomElement node, QTextTable *table) /*нерекурсивная функция обхода дерева таблицы для формата odt. В цикле ищется
1 строка, затем ее 1 элемент, находится текст ячейки, он обрабатывается с помощью вспомогательного метода, и затем вносится в текущую ячейку,
и так пока все строчки не закончатся*/
{
    QDomElement row = node.firstChildElement();
    QDomElement cell;
    QDomElement text;
    QString id;
    QString id2;
    int i = 0;
    int j = 0;
    QTextTableCell cellValue;
    QTextCursor cellCursor;
    QTextCharFormat charFormat;

    while(!row.isNull())
    {
        if(row.tagName() == "table:table-row")
        {
             j = 0;
             cell = row.firstChildElement();
             while(!cell.isNull())
             {
                if(cell.tagName() == "table:table-cell")
                {
                    cellValue = table->cellAt(i,j);
                    cellCursor = cellValue.firstCursorPosition();
                    text = cell.firstChildElement();

                    id = text.attribute("text:style-name");
                    id2 = text.firstChildElement().attribute("text:style-name");
                    GetFontsODT(node, id, id2, charFormat);

                    if (cell.isNull()){
                        cellCursor.insertText(" ");
                    }else {
                        cellCursor.insertText(cell.text(), charFormat);
                    }
                }
                ++j;
                cell = cell.nextSiblingElement();
             }
             ++i;
        }
        row = row.nextSiblingElement();
    }
}

int myclass::GetRowsODT(QDomElement node) /*вспомогательная функция для поиска количества строк в таблице для формата odt*/
{
    QDomElement child = node.firstChildElement();
    int rows = 0;
    while(!child.isNull())
    {
        if(child.tagName() == "table:table-row")
        {
             ++rows;
        }
        child = child.nextSiblingElement();
    }
    return rows;
}

int myclass::GetColumnsODT(QDomElement node) /*вспомогательная функция для поиска количества столбцов в таблице для формата odt*/
{
    QDomElement child = node.firstChildElement();
    QDomElement row;
    int maxCols = 0;
    int cols = 0;
    while(!child.isNull())
    {
        cols = 0;
        if(child.tagName() == "table:table-row")
        {
             row = child.firstChildElement();
             while(!row.isNull())
             {
                if(row.tagName() == "table:table-cell")
                {
                    ++cols;
                }
                row = row.nextSiblingElement();
             }
        }
        if(maxCols < cols)
        {
             maxCols = cols;
        }
        child = child.nextSiblingElement();
    }
    return maxCols;
}

void myclass::GetFontsODT(QDomElement node, QString id, QString id2, QTextCharFormat &format) /*вспомогательная функция для форматирования текста
каждой ячейки таблицы*/
{
    QFont font;
    QDomNode parent = node.parentNode().parentNode();
    QString parentId;
    QColor textColor;
    QColor backColor;
    textColor.setNamedColor("#000");
    backColor.setNamedColor("#fff");

    while(!parent.isNull()) // поиск office:automatic-styles
    {
        if(parent.toElement().tagName() == "office:automatic-styles")
        {
             break;
        }
        parent = parent.previousSibling();
    }
    QDomElement style = parent.firstChildElement();
    QDomElement style2;
    while(!style.isNull())
    {
        if(style.attribute("style:name") == id)   //поиск стиля с нужным именем атрибута
        {
             parentId = style.attribute("style:parent-style-name");
             style = style.firstChildElement();
        }
        if(style.tagName() == "style:text-properties") //переход к характеристикам текста
        {
             break;
        }
        style = style.nextSiblingElement();
    }

    if(style.tagName() != "style:text-properties")  //если стиль не был найден, то будет использован стиль из файла styles.xml
    {
        style = parent.firstChildElement();
        while(!style.isNull())
        {
             if(style.attribute("style:name") == id2)
             {
                style = style.firstChildElement();
             }
             if(style.tagName() == "style:text-properties")
             {
                break;
             }
             style = style.nextSiblingElement();
        }

        style2 = stylesInfo.firstChildElement().firstChildElement();
        style2 = style2.nextSiblingElement().firstChildElement();
        while(!style2.isNull())
        {
             if(style2.attribute("style:name") == id)
             {
                style2 = style2.firstChildElement();
                break; // FIND
             }
             style2 = style2.nextSiblingElement();
        }

        while(!style2.isNull())
        {
             if(style2.tagName() == "style:text-properties")
             {
                break;
             }
             style2 = style2.nextSiblingElement();
        }
    }
    else
    {
        style2 = stylesInfo.firstChildElement().firstChildElement();
        style2 = style2.nextSiblingElement().firstChildElement();
        while(!style2.isNull())
        {
             if(style2.attribute("style:name") == parentId)
             {
                style2 = style2.firstChildElement();
                break;
             }
             style2 = style2.nextSiblingElement();
        }

        while(!style2.isNull())
        {
             if(style2.tagName() == "style:text-properties")
             {
                break;
             }
             style2 = style2.nextSiblingElement();
        }
    }

    if(style.attribute("style:font-name") == "")  //обработка шрифта
    {
        font.setFamily(style2.attribute("style:font-name"));
    }
    else
    {
        font.setFamily(style.attribute("style:font-name"));
    }

    if(style.attribute("fo:font-size") == "")   // обработка размера шрифта
    {
        QString str_font = style2.attribute("fo:font-size");
        str_font.replace("pt","");
        font.setPointSize(str_font.toDouble()*1.5);
    }
    else
    {
        QString str_font = style.attribute("fo:font-size");
        str_font.replace("pt","");
        font.setPointSize(str_font.toDouble()*1.5);
    }

    if(style.attribute("fo:font-style") == "")  //обработка курсива
    {
        if(style2.attribute("fo:font-style") == "italic")
        {
            font.setItalic(true);
        }
    }
    else
    {
        if(style.attribute("fo:font-style") == "italic")
        {
            font.setItalic(true);
        }
    }

    if(style.attribute("fo:font-weight") == "")  //обработка жирного
    {
        if(style2.attribute("fo:font-weight") == "bold")
        {
            font.setBold(true);
        }
    }
    else
    {
        if(style.attribute("fo:font-weight") == "bold")
        {
            font.setBold(true);
        }
    }

    if(style.attribute("fo:color").isNull())   // обработка цвета текста
    {
        if(!style2.attribute("fo:color").isNull())
        {
            textColor.setNamedColor(style2.attribute("fo:color"));

        }
    }
    else
    {
        if(!style.attribute("fo:color").isNull())
        {

            textColor.setNamedColor(style.attribute("fo:color"));

        }
    }
    format.setForeground(textColor);
    format.setFont(font);
}

void myclass::SearchTagDOCX(QDomElement node, paragNode *var_node) // поиск тега и заполнение структуры paragNode в зависимости от него. Функция switch-case
{
    int tag_num = GetTegNumber(node.toElement().tagName());
    switch (tag_num) {
    case 0: // "w:t" - set text
        if (node.text().isEmpty()){
             var_node->text = " ";
        }else {
             var_node->text = node.text();
        }
        break;
    case 1: // "w:rFonts" - set font
        var_node->font.setFamily(node.attribute("w:ascii"));
        break;
    case 2: // "w:b" - set bolding
        if (node.attribute("w:val") != "false" && (node.parentNode().parentNode().toElement().tagName() == "w:r" ||
                                                   node.parentNode().parentNode().toElement().tagName() == "w:style")){
             var_node->font.setBold(true);
        }
        break;
    case 3: // "w:i" - set italic
        if (node.attribute("w:val") != "false" && (node.parentNode().parentNode().toElement().tagName() == "w:r" ||
                                                   node.parentNode().parentNode().toElement().tagName() == "w:style")){
             var_node->font.setItalic(true);
        }
        break;
    case 4:
        if (node.attribute("w:val") != "none"){
             var_node->font.setUnderline(true);
        }
        break;
    case 5: // "w:pStyle" - set title & headings
        SetStyleDOCX(node.attribute("w:val"), var_node);
        break;
    case 6: // "w:sz"
        var_node->font.setPixelSize(node.attribute("w:val").toDouble());
        break;
    case 7: // "w:color" - set
        bool ok;
        var_node->colorText = node.attribute("w:val").toInt(&ok,16);
        break;
    case 8: // "w:highlight
        var_node->colorBackground = node.attribute("w:val");
        break;
    case 9: // "w:spacing" - set margin
        if (node.hasAttribute("w:before")){
             var_node->indentParag.top = node.attribute("w:before").toInt()/10;
        }
        if (node.hasAttribute("w:after")){
             var_node->indentParag.bottom = node.attribute("w:after").toInt()/10;
        }
        break;
    case 10: // "w:jc" - align
        if (node.attribute("w:val") == "center"){
            var_node->alignParag = Qt::AlignCenter;
        }
        if (node.attribute("w:val") == "left"){
            var_node->alignParag = Qt::AlignLeft;
        }
        if (node.attribute("w:val") == "right"){
            var_node->alignParag = Qt::AlignRight;
        }
        if (node.attribute("w:val") == "both"){
            var_node->alignParag = Qt::AlignJustify;
        }
        break;
    case 11: // "w:ind" - indents of paragraph
        if (!node.attribute("w:left").isEmpty()){
             var_node->indentParag.left += (int)(node.attribute("w:left").toInt()*var_node->indentParag.proport);
        }
        if (!node.attribute("w:right").isEmpty()){
             var_node->indentParag.right += (int)(node.attribute("w:right").toInt()*var_node->indentParag.proport);
        }
        if (!node.attribute("w:firstLine").isEmpty()){
             var_node->indentParag.firstLine = (int)(node.attribute("w:firstLine").toInt()*var_node->indentParag.proport);
        }
        if (!node.attribute("w:hanging").isEmpty()){
             var_node->indentParag.firstLine = -(int)(node.attribute("w:hanging").toInt()*var_node->indentParag.proport);
        }
        break;
    case 12: // "w:p" - parag or not parag
        var_node->newParag = true;
        break;
    case 13: // "w:blip"
        var_node->image_id = node.attribute("r:embed");
        break;
    case 14:  // "w:ilvl"
        var_node->isPartOfList = true;
        if (node.parentNode().toElement().tagName() == "w:numPr"){
            GetNumberIdDOCX(node.attribute("w:val"), node.nextSiblingElement().attribute("w:val"), var_node);
        }
        break;
    case 15: // "w:numFmt"
        var_node->listStyle = node.attribute("w:val");
        if (var_node->listStyle == "bullet"){
            ui->textEdit->append("●\t");
        }
        break;
    case 16: // "w:lvlJc"
        if (node.attribute("w:val") == "center"){
            var_node->alignParag = Qt::AlignCenter;
        }
        if (node.attribute("w:val") == "left"){
            var_node->alignParag = Qt::AlignLeft;
        }
        if (node.attribute("w:val") == "right"){
            var_node->alignParag = Qt::AlignRight;
        }
        if (node.attribute("w:val") == "both"){
            var_node->alignParag = Qt::AlignJustify;
        }
        break;
    default:
        break;
    }
}

int myclass::GetTegNumber(QString tag_name) // вспомагательная функция для определения номера тега в списке тегов TextTagsNameQList
{
    for (int i = 0; i < TextTagsNameQList.size(); i++){
        if (tag_name == TextTagsNameQList[i]){
            return i;
        }
    }
    return -1;
}

void myclass::SetStyleDOCX(QString attribute, paragNode *var_node) /* вспомагательная функция для поиска стиля шрифта в отдельном документе.
 Работает  по аналолгии с GetDataFromDOCX, но рекурсия заменена циклом (полных обход дерева не требуется, нужна только ветка для узла с
искомым стилем*/
{
    QDomElement root = stylesInfo.firstChildElement();
    QDomElement node = root.firstChildElement();

    while (!node.isNull())
    {
        if (node.attribute("w:styleId") == attribute){ /* если у узла атрибут w:styleId совпадает с имеющимся у нас стилем attribute, то заходим в его поддерево
        иначе переходим к соседнему узлу*/
            node = node.firstChildElement();
            break;
        }
        node = node.nextSiblingElement();
    }

    // ищем нужные нам теги
    while (!node.isNull())
    {
        if (node.tagName() == "w:pPr"){
            node = node.firstChildElement();
            while(!node.isNull())
            {
                SearchTagDOCX(node, var_node);
                if (node.nextSiblingElement().isNull()){
                    node = node.parentNode().toElement();
                    break;
                }else {
                    node = node.nextSiblingElement();
                }
            }
        }
        if (node.tagName() == "w:rPr"){
            node = node.firstChildElement();
            while(!node.isNull())
            {
                SearchTagDOCX(node, var_node);
                if (node.nextSiblingElement().isNull()){
                    node = node.parentNode().toElement();
                    break;
                }else {
                    node = node.nextSiblingElement();
                }
            }
        }
        node = node.nextSiblingElement();
    }
}

void myclass::GetNumberIdDOCX(QString attr_ilvl, QString attr_numId, paragNode *var_node){ /* вспомогательная функция для SetNumberingDOCX
для определения abstractNumId (abstractNumId в файле numbering.xml и numId в файле document.xml различны, их соотношение хранится вконце файла
numbering.xml)*/
    QDomElement root = listInfo.firstChildElement();
    QDomElement node = root.firstChildElement();
    while (!node.isNull())
    {
        if (node.tagName() == "w:num" && node.attribute("w:numId") == attr_numId){
            attr_numId = node.firstChildElement().attribute("w:val");
            break;
        }
        node = node.nextSiblingElement();
    }
    SetNumberingDOCX(attr_ilvl, attr_numId, var_node);
}

void myclass::SetNumberingDOCX(QString attr_ilvl, QString attr_numId, paragNode *var_node) /*вспомагательная функция для поиска стиля спписка в отдельном документе.
Работает  по аналолгии с SetStyleDOCX*/
{
    QDomElement root = listInfo.firstChildElement();
    QDomElement node = root.firstChildElement();

    while (!node.isNull())
    {
        if (node.attribute("w:abstractNumId") == attr_numId){
            node = node.firstChildElement();
            break;
        }
        node = node.nextSiblingElement();
    }

    while (!node.isNull())
    {
        if (node.tagName() == "w:lvl" && node.attribute("w:ilvl") == attr_ilvl){
            node = node.firstChildElement();
            while(!node.isNull())
            {
                SearchTagDOCX(node, var_node);
                if (node.nextSiblingElement().isNull()){
                    node = node.parentNode().toElement();
                    break;
                }else {
                    node = node.nextSiblingElement();
                }
                if (node.tagName() == "w:pPr"){
                    SearchTagDOCX(node.firstChildElement(), var_node);
                }
            }
        }
        node = node.nextSiblingElement();
    }
}

void myclass::ClearNode(paragNode *var_node)
{
    var_node->text = "";
    var_node->image_id = "";
    var_node->font_id = "";
    var_node->style_id = "";
    var_node->line_style = "";
    var_node->font.setUnderline(false);
    var_node->font.setBold(false);
    var_node->font.setItalic(false);
    var_node->font.setFamily("Calibri");
    var_node->font.setPixelSize(24);
    var_node->newParag = false;
    var_node->alignParag= Qt::AlignLeft;
    var_node->indentParag.left = 170;
    var_node->indentParag.right = 85;
    var_node->indentParag.firstLine = 0;
    var_node->indentParag.top = 0;
    var_node->indentParag.bottom = 0;
    var_node->colorText = 0;
    var_node->colorBackground = "white";
    var_node->brush_back.setNamedColor("#fff");
    var_node->brush_char.setNamedColor("#000");
    var_node->listStyle = "";
    var_node->isPartOfList = false;
}

void myclass::GetDataFromODT(QDomElement node, paragNode *var_node)/* рекурсивная функция для обхода дерева и вывода отформатированного
    текста. Реализован обход на основе прямого. Сначала идет функция идет вниз по дереву по child-элементам, пока не окажется в листе (первая
    рекурсия закончилась). После этого возвращается в пройденные узлы и, в зависимости от тега узла, сохраняет заполненную структуру paragNode (с
    текстом и форматированием) либо выводит на экран отформатированный текст. После чего осуществляется второй рекурсивный переход к
    соседнему (по отношению к текущему) узлу. Работает по аналогии с GetDataFromDOCX()*/
{

    if (!node.toElement().isNull()){
        if (node.toElement().tagName() == "table:table"){
            if (!node.nextSiblingElement().isNull()){
                int rows = GetRowsODT(node);
                int columns = GetColumnsODT(node);
                QTextCursor cursor2(ui->textEdit->textCursor());
                QTextTableFormat format;
                format.setBorderCollapse(true);
                format.setCellPadding(15);
                format.setTopMargin(10);
                format.setBottomMargin(10);
                format.setLeftMargin(170);
                format.setRightMargin(170);
                QTextTable *table = cursor2.insertTable(rows, columns, format);
                ParsTableODT(node, table);

                GetDataFromODT(node.nextSiblingElement(), var_node);
            }
            return;
        }
        //исключает узлы, содержащие стили текста и абзаца при обходе
        if(node.tagName()!="style:text-properties" && node.tagName()!="style:paragraph-properties"){
            //
            SearchTagOdt(node, var_node);// заполнение экземпляра структуры paragNode в зависимости от тега
        }
        if (!node.firstChildElement().isNull()){
            GetDataFromODT(node.firstChildElement(), var_node);// первый рекурсивный переход к child-элементу
        }
        if (node.tagName() == "text:span"||node.tagName() == "text:p"||node.tagName() == "text:h"){
            //запоминание форматирования для абзаца, заголовка и отдельных частей абзаца
            var_node->alignParag = old_var_node.alignParag;
            paragComponentsList.push_back(*var_node);
            if(node.tagName() == "text:p"){//отчистка форматирования абзаца
                ClearNode(&old_var_node);
            }
            ClearNode(var_node);//очистка форматирования для абзаца, заголовка и отдельных частей абзаца

        }
        if((node.tagName() == "text:p"||node.tagName() == "text:h") && !paragComponentsList.isEmpty()){// случай, когда в абзаце или заголовке есть текст
            while(!paragComponentsList.isEmpty()){

                *var_node = paragComponentsList.front();
                paragComponentsList.pop_front();
                QTextCursor cursor = QTextCursor(ui->textEdit->textCursor());
                QTextBlockFormat blockFormat = cursor.blockFormat();
                // установака отступов, межстрочных интервалов, отступа абзаца и выравнивания для абзаца
                if (var_node->listStyle.isEmpty()){
                    blockFormat.setLeftMargin(var_node->indentParag.left);
                } else {
                    blockFormat.setLeftMargin(var_node->indentParag.left - 1000);
                }
                blockFormat.setRightMargin(var_node->indentParag.right);
                blockFormat.setTopMargin(var_node->indentParag.top);
                blockFormat.setBottomMargin(var_node->indentParag.bottom);
                blockFormat.setTextIndent(var_node->indentParag.firstLine);
                blockFormat.setAlignment(var_node->alignParag);
                cursor.setBlockFormat(blockFormat);
                // установка форматирования текста
                QTextCharFormat charFormat;
                charFormat.setFont(var_node->font);
                if(var_node->line_style!=""){
                    SetLineStyleOdt(var_node,charFormat);
                }
                charFormat.setBackground(var_node->brush_back);
                charFormat.setForeground(var_node->brush_char);
                cursor.insertText(var_node->text, charFormat);

                if (!var_node->image_id.isEmpty()){ // вывод изображения
                    PrintImageOdt(var_node);
                }
            }
            ui->textEdit->append("");
            ClearNode(var_node);

        }
        if (!node.nextSiblingElement().isNull()){
            GetDataFromODT(node.nextSiblingElement(), var_node);// второй рекурсивный переход к соседнему (по отношению к текущему) узлу
        }
    }
}

void myclass::SearchTagOdt(QDomElement node, paragNode *var_node){/*метод, в зависимости от тега,
заполняет экземпляр структуры paragNode*/
    int tag_num = GetTegNumber(node.toElement().tagName());
    QString str_font;
    switch (tag_num) {
    case 0://text:p
        if (node.parentNode().toElement().tagName() == "text:list-item"){
            var_node->text = var_node->listSymb.dot + "\t" + node.text();
        }else {
            if(node.firstChildElement().isNull()){
                if (node.text().isEmpty()){
                    var_node->text = " ";
                }
                else {
                    var_node->text = node.text();
                }
            }
        }
        var_node->font_id=node.attribute("text:style-name");
        SetStyleOdt(var_node);
        break;
    case 1://text:h
        var_node->text = node.text();
        var_node->font_id = node.attribute("text:style-name");
        var_node->style_id = node.attribute("text:style-name");
        SetStyleOdt(var_node);
        break;
    case 2://text:span
        if(node.firstChildElement().tagName() != "draw:frame"){
            if (node.text().isEmpty()){
                var_node->text = " ";
            }
            else {
                var_node->text = node.text();
                var_node->font_id=node.attribute("text:style-name");
                SetStyleOdt(var_node);
            }
        }
        break;
    case 3://draw:image
        var_node->image_id = node.attribute("xlink:href");
        break;
    case 4://style:text-properties
        var_node->font.setFamily(node.attribute("style:font-name"));
        str_font = node.attribute("fo:font-size");
        str_font.replace("pt","");
        var_node->font.setPointSize(str_font.toDouble()*1.5);
        if(node.attribute("fo:font-style")=="italic"){
            var_node->font.setItalic(true);
        }
        if(node.attribute("fo:font-weight")=="bold"){
            var_node->font.setBold(true);
        }
        if(!node.attribute("fo:background-color").isNull()){
            var_node->brush_back.setNamedColor(node.attribute("fo:background-color"));
        }
        if(!node.attribute("fo:color").isNull()){
            var_node->brush_char.setNamedColor(node.attribute("fo:color"));
        }
        break;
    default:
        break;
    }
}
/*метод задает стиль подчеркиваний для текста*/
void myclass::SetLineStyleOdt(paragNode *var_node, QTextCharFormat &charFormat){
    if(var_node->line_style=="single"){
        charFormat.setUnderlineStyle(QTextCharFormat::SingleUnderline);
    }
    else if(var_node->line_style=="double"){
        charFormat.setUnderlineStyle(QTextCharFormat::SingleUnderline);
    }
    else if(var_node->line_style=="thick"){
        charFormat.setUnderlineStyle(QTextCharFormat::SingleUnderline);
    }
    else if(var_node->line_style=="wave"){
        charFormat.setUnderlineStyle(QTextCharFormat::WaveUnderline);
    }
    else if(var_node->line_style=="dotted"){
        charFormat.setUnderlineStyle(QTextCharFormat::DotLine);
    }
    else if(var_node->line_style=="dash"){
        charFormat.setUnderlineStyle(QTextCharFormat::DashUnderline);
    }
    else if(var_node->line_style=="dot-dash"){
        charFormat.setUnderlineStyle(QTextCharFormat::DashDotLine);
    }
    else if(var_node->line_style=="dot-dot-dash"){
        charFormat.setUnderlineStyle(QTextCharFormat::DashDotDotLine);
    }
}

void myclass::SetStyleOdt(paragNode *var_node){/*вспомогательный метод для поиска узла,
содержащего стили для текста в content.xml. Поиск производится по атрибуту "style:name" тега "style:style" */
    QDomElement root = document.firstChildElement();
    QDomElement node = root.firstChildElement();
    node = node.nextSiblingElement().firstChildElement();
    FindStyleNameOdt(node,var_node);

    if(node.tagName()!="style:style"){
        if(FindPropertiesOdt(node,var_node)==true){
            return;
        }

    }
    SetStyleFromStylesOdt(var_node);
}

void myclass::SetStyleFromStylesOdt(paragNode *var_node){/*вспомогательный метод для поиска узла,
содержащего стили текста в style.xml. Вызывается, если в content.xml не был найден тег style:text-properties*/
    QDomElement root = stylesInfo.firstChildElement();
    QDomElement node = root.firstChildElement();
    node = node.nextSiblingElement().firstChildElement();

    FindStyleNameOdt(node,var_node,true);
    FindPropertiesOdt(node,var_node);
}

void myclass::FindStyleNameOdt(QDomElement &node, paragNode *var_node, bool flag){/*вспомогательный метод для поиска узла
    с соответствующим идентификатором стилей*/

    while (!node.isNull()){
        if (node.attribute("style:name") == var_node->font_id && flag == false){
            var_node->style_id = node.attribute("style:parent-style-name");
            if(!node.firstChildElement().isNull()){
                node = node.firstChildElement();
            }
            break;
        }
        else if(node.attribute("style:name")==var_node->style_id && flag == true){
            if(!node.firstChildElement().isNull()){
                node = node.firstChildElement();
            }
            break;

        }
        node = node.nextSiblingElement();
    }
}

bool myclass::FindPropertiesOdt(QDomElement &node, paragNode *var_node){/*вспомогательный метод для поиска узлов,
содержащих в себе свойства абзаца и свойства текста.
Из этого метода вызывается Switch-case метод, задающий соответсвующие стили текущему узлу*/
    while(!node.isNull()){
        if(node.tagName()=="style:paragraph-properties"){
            SearchTagOdt(node,var_node);

        }
        if(node.tagName()=="style:text-properties"){
            SearchTagOdt(node,var_node);
            return true;
        }
        node = node.nextSiblingElement();
    }
    return false;
}

void myclass::PrintImageDocx(paragNode *var_node){/*вставка изображения для docx*/

    QDomElement root = imageName.firstChildElement();
    QDomElement node = root.firstChild().toElement();
    QString imagePath;
    while (!node.isNull())//поиск пути к изображению в document.xml.rels
    {
        if(node.attribute("Id") == var_node->image_id){
            imagePath = node.attribute("Target");
            break;
        }
        node = node.nextSiblingElement();
    }

    QZipReader zip(path, QIODevice::ReadOnly);

    zip.extractAll(path);
    if (zip.status() == QZipReader::NoError)
    {
        QImage img;
        QByteArray data = zip.fileData("word/"+imagePath);

        img.loadFromData(data);

        QTextCursor cursor = QTextCursor(ui->textEdit->textCursor());
        cursor.insertImage(img);


        zip.close();

    }
    else
    {
        qDebug() << "Faild zip";
        QMessageBox::critical(nullptr,"Error::File not found","Check the file path.\nFile does not exist.");
    }
}

void myclass::PrintImageOdt(paragNode *var_node){/*вставка изображения для odt*/
    QZipReader zip(path, QIODevice::ReadOnly);

    zip.extractAll(path);
    if (zip.status() == QZipReader::NoError)
    {
        QImage img;
        QByteArray data = zip.fileData(var_node->image_id);
        img.loadFromData(data);
        QTextCursor cursor = QTextCursor(ui->textEdit->textCursor());
        cursor.insertImage(img);
        zip.close();

    }
    else
    {
        qDebug() << "Faild zip";
        QMessageBox::critical(nullptr,"Error::File not found","Check the file path.\nFile does not exist.");
    }
}
