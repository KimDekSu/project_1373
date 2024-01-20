#include "myclass.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);//создание приложения с графическим интерфейсом
    QCommandLineParser parser;//для обработки командной строки
    parser.addPositionalArgument("file", "File to load");//определяет дополнительный аргумент приложения
    parser.process(app);//обработка командной строки, полученной из QApplication
    const QStringList args = parser.positionalArguments();//получение списка позиционных аргументов
    myclass w;//класс приложения
    if (args.isEmpty()) {
        //если не передан путь к файлу в качестве входного аргумента
        w.show();//показать окно приложения
    } else {
        //если передан путь к файлу в качестве входного аргумента
        QString path = args.first();//возвращает первый элемент из QStringList
        if(w.setPath(path)==false){//если не удалось открыть файл
            return app.exec();//закрываем приложение
        }
        w.show();//показать окно приложения
    }
    return app.exec();//закрыть окно приложения
}
