#ifndef MYCLASS_H
#define MYCLASS_H

#include <QMainWindow>
#include <QApplication>
#include <QMessageBox>

#include <QDebug>
#include <qstring.h>
#include <qtextstream.h>
#include <QFileDialog>
#include <QtXml>

#include <qzipreader_p.h>
#include <qzipwriter_p.h>

#include <QPushButton>
#include <QTextBlock>

#include <QImage>
#include <QTextCursor>
#include <QImageReader>
#include <QTextTableCell>

QT_BEGIN_NAMESPACE
namespace Ui { class myclass; }
QT_END_NAMESPACE

class myclass : public QMainWindow
{
    Q_OBJECT

public:
    myclass(QWidget *parent = nullptr);
    ~myclass();
    bool setPath(QString path);
private:
    Ui::myclass *ui;
    struct paragNode // вся информация о части текста
    {
        bool newParag = false;
        bool isPartOfList = false;

        QString text; // текст
        QFont font;// стиль текста
        int colorText = 0;
        QString colorBackground = "#ffffff";
        struct
        {
            double proport = 0.16;
            int left = (int)(1701*proport);
            int right = (int)(850*proport);
            int firstLine;
            int top = 0;
            int bottom = 0;
        } indentParag;
        Qt::AlignmentFlag alignParag = Qt::AlignLeft;

        // odt
        QString font_id;// идентификатор тега для поиска стилей в content.xml
        QString style_id;// идентификатор тега для поиска стилей в style.xml
        QString line_style;// стиль подчеркиваний текста
        QColor brush_back;// цвет фона текста
        QColor brush_char;// цвет текста

        // идентификатор изображения
        QString image_id;

        // для списка
        QString listStyle = "";
        struct{
            QString dot = "●";
            int num = 1;
        } listSymb;
    };
    paragNode old_var_node; // параметры форматирования для всего абзаца
    QList<paragNode> paragComponentsList; // список для запоминания частей абзаца с их форматированием
    QList<QString> TextTagsNameQList; // список нужных нам тегов
    QDomDocument stylesInfo; // вспомогательный документ со стилями
    QDomDocument imageName; // вспомогательный документ с изобрадением
    QDomDocument document; // основвной документ
    QDomDocument listInfo; // вспомогательный документ со списками
    QString path; // путь к файлу

private slots:
    void on_actionOpen_File_triggered();
    bool OpenXML(QString path); // открытие файлов
    void ParsXML(QDomDocument document, QString file_exten); /* заполнение списка тегов, объявление и инициализация paragNode, запуск рекурсии
    GetDataFromDOCX*/
    void GetDataFromDOCX(QDomElement node, paragNode *var_node); /* рекурсивная функция для обхода дерева и вывода отформатированного
    текста. Реализован обход на основе прямого. Сначала идет функция идет вниз по дереву по child-элементам, пока не окажется в листе (первая
    рекурсия закончилась). После этого возвращается в пройденные узлы и, в зависимости от тега узла, сохраняет заполненную структуру paragNode (с
    текстом и форматированием) либо выводит на экран отформатированный текст. После чего осуществляется второй рекурсивный переход к
    соседнему (по отношению к текущему) узлу.*/
    void SearchTagDOCX(QDomElement node, paragNode *var_node); // поиск тега и заполнение структуры paragNode в зависимости от него. Функция switch-case
    void SetStyleDOCX(QString attribute, paragNode *var_node); /* вспомагательная функция для поиска стиля шрифта в отдельном документе.
    Работает  по аналолгии с GetDataFromDOCX, но рекурсия заменена циклом (полных обход дерева не требуется, нужна только ветка для узла с
    искомым стилем*/
    void GetNumberIdDOCX(QString attr_ilvl, QString attr_numId, paragNode *var_node); /* вспомогательная функция для SetNumberingDOCX
    для определения abstractNumId (abstractNumId в файле numbering.xml и numId в файле document.xml различны, их соотношение хранится вконце файла
    numbering.xml)*/
    void SetNumberingDOCX(QString attr_numId, QString attr_ilvl, paragNode *var_node);/*вспомагательная функция для поиска стиля спписка в отдельном документе.
    Работает  по аналолгии с SetStyleDOCX*/

    int GetTegNumber(QString tag_name); // вспомагательная функция для определения номера тега в списке тегов TextTagsNameQList
    void ClearNode(paragNode *var_node); // очистка ноды

    void GetDataFromODT(QDomElement node, paragNode *var_node);/* рекурсивная функция для обхода дерева и вывода отформатированного
    текста. Реализован обход на основе прямого. Сначала идет функция идет вниз по дереву по child-элементам, пока не окажется в листе (первая
    рекурсия закончилась). После этого возвращается в пройденные узлы и, в зависимости от тега узла, сохраняет заполненную структуру paragNode (с
    текстом и форматированием) либо выводит на экран отформатированный текст. После чего осуществляется второй рекурсивный переход к
    соседнему (по отношению к текущему) узлу.*/
    void SetLineStyleOdt(paragNode *var_node, QTextCharFormat &charFormat);/*метод задает стиль подчеркиваний для текста*/
    void SetStyleOdt(paragNode *var_node);/*вспомогательный метод для поиска узла, содержащего стили для текста в content.xml*/
    void SetStyleFromStylesOdt(paragNode *var_node);/*вспомогательный метод для поиска узла, содержащего стили текста в style.xml*/
    void FindStyleNameOdt(QDomElement &node, paragNode *var_node, bool flag = false);/*вспомогательный метод для поиска узла
    с соответствующим идентификатором стилей*/
    bool FindPropertiesOdt(QDomElement &node, paragNode *var_node);/*вспомогательный метод для поиска узлов, содержащих в себе свойства абзаца и свойства текста.
    Из этого метода вызывается Switch-case метод, задающий соответсвующие стили текущему узлу*/
    void SearchTagOdt(QDomElement node, paragNode *var_node);/*метод, в зависимости от тега, заполняет экземпляр структуры paragNode*/

    void PrintImageDocx(paragNode *var_node);/*вставка изображения для docx*/
    void PrintImageOdt(paragNode *var_node);/*вставка изображения для odt*/

    int GetRowsDOCX(QDomElement node); /*получение количества строк таблицы для docx*/
    int GetColumnsDOCX(QDomElement node); /*получение количества столбцов таблицы для docx*/
    void ParsTableDOCX(QDomElement node, QTextTable *table, int rows, int columns); /*нерекурсивная функция для прохода дерева таблицы с заполнением таблицы
    по ячейно и последующим форматированием текста каждой ячейки для docx*/
    int GetColumnsODT(QDomElement node); /*получение количества столбцов таблицы для odt*/
    int GetRowsODT(QDomElement node); /*получение количества строк таблицы для odt*/
    void ParsTableODT(QDomElement node, QTextTable *table); /*нерекурсивная функция для прохода дерева таблицы с заполнением таблицы
    по ячейно и последующим форматированием текста каждой ячейки при помощи вспомогательного метода для odt*/
    void GetFontsODT(QDomElement node, QString id, QString id2, QTextCharFormat &format); /*вспомогательный метод для форматирования текста табличных ячеек,
    ищет и применяет к формату табличной ячейки шрифт, его размер, жирный или нет, курсив или нет и т.п.*/

};
#endif // myclass_H
