#include "calc_salary_birulevo.h"

// ============================================================
// Вспомогательные функции
// ============================================================

// Читает число из ячейки
// Обрабатывает: 2,00 | 50 000,00 | 1234.56 | формулы с кэшем
static double readCellDouble(QXlsx::Document &doc, int row, int col) {
    auto cell = doc.cellAt(row, col);
    if (!cell)
        return 0.0;

    QVariant val = cell->value();
    if (!val.isValid() || val.isNull())
        return 0.0;

    // Если уже число — сразу возвращаем
    if (val.typeId() == QMetaType::Double ||
        val.typeId() == QMetaType::Float ||
        val.typeId() == QMetaType::Int ||
        val.typeId() == QMetaType::LongLong) {
        return val.toDouble();
    }

    // Если строка — парсим
    if (val.canConvert<QString>()) {
        QString str = val.toString().trimmed();

        // Если это формула без кэша
        if (str.startsWith("=")) {
            qWarning() << "Формула без кэша:" << row << col << str;
            return 0.0;
        }

        // Убираем все виды пробелов (обычный, неразрывный, тонкий)
        str.remove(' ');
        str.remove(QChar(0x00A0)); // неразрывный пробел
        str.remove(QChar(0x202F)); // узкий неразрывный пробел
        str.remove(QChar(0x2009)); // тонкий пробел

        // Заменяем запятую на точку: "2,00" -> "2.00"
        str.replace(',', '.');

        // Убираем символ валюты если есть
        str.remove(QChar(0x20BD)); // ₽ (символ рубля)
        str.remove(QString::fromUtf8("р"));
        str.remove(QChar('$'));

        str = str.trimmed();

        bool ok = false;
        double d = str.toDouble(&ok);
        if (ok) return d;

        qDebug() << "Не удалось распарсить число:" << val.toString()
                 << "-> после обработки:" << str;
    }

    return 0.0;
}

// Читает строку из ячейки (как есть)
static QString readCellString(QXlsx::Document &doc, int row, int col) {
    auto cell = doc.cellAt(row, col);
    if (!cell)
        return QString();

    QVariant val = cell->value();
    if (!val.isValid() || val.isNull())
        return QString();

    return val.toString().trimmed();
}

// Сравнение фамилий без учёта регистра и пробелов
static bool namesMatch(const QString &a, const QString &b) {
    return a.trimmed().compare(b.trimmed(), Qt::CaseInsensitive) == 0;
}

// ============================================================
// Основная функция
// ============================================================

QVariantMap CalcBirulevo::execute(const QVariantMap &params)
{
    QVariantMap result;

    QStringList files = params.value("xlsxFiles").toStringList();
    QString outputFolder = params.value("outputFolder").toString();
    QString outputFile = params.value("outputFile").toString();

    if (files.isEmpty()) {
        result["success"] = false;
        result["error"] = "Не переданы XLSX-файлы";
        return result;
    }

    if (outputFolder.isEmpty() || !QDir(outputFolder).exists()) {
        result["success"] = false;
        result["error"] = "Не указана или не существует папка для результатов";
        return result;
    }

    QXlsx::Document outputDoc;
    QList<QString> surnames;
    QMap<QString, int> surnameNextRow;

    // ============ Стили ============
    QXlsx::Format headerFormat;
    headerFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    headerFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    headerFormat.setFontBold(true);
    headerFormat.setFontSize(14);
    headerFormat.setPatternBackgroundColor(QColor(173, 216, 230));
    headerFormat.setFontColor(QColor(0, 0, 0));
    headerFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    headerFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    QXlsx::Format colHeaderFormat;
    colHeaderFormat.setFontBold(true);
    colHeaderFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    colHeaderFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    colHeaderFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    colHeaderFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    QXlsx::Format dataFormat;
    dataFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    dataFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);

    QXlsx::Format sumFormat;
    sumFormat.setFontBold(true);
    sumFormat.setHorizontalAlignment(QXlsx::Format::AlignHCenter);
    sumFormat.setVerticalAlignment(QXlsx::Format::AlignVCenter);
    sumFormat.setPatternBackgroundColor(QColor(255, 255, 0));
    sumFormat.setFontColor(QColor(0, 0, 0));
    sumFormat.setBottomBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setTopBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setLeftBorderStyle(QXlsx::Format::BorderMedium);
    sumFormat.setRightBorderStyle(QXlsx::Format::BorderMedium);

    QXlsx::Format headerFillFormat;
    headerFillFormat.setPatternBackgroundColor(QColor(173, 216, 230));
    headerFillFormat.setBottomBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setTopBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setLeftBorderStyle(QXlsx::Format::BorderThin);
    headerFillFormat.setRightBorderStyle(QXlsx::Format::BorderThin);

    // ============ Обработка файлов ============
    for (const QString &filepath : files) {
        QXlsx::Document xlsx(filepath);
        if (!xlsx.load()) {
            qWarning() << "Не удалось открыть:" << filepath;
            continue;
        }

        // Дата из E1
        double dateVal = readCellDouble(xlsx, 1, 5);
        int monthNum = qRound((dateVal - (int)dateVal) * 100);

        // Год из I1
        QString year = readCellString(xlsx, 1, 9);
        if (year.endsWith("г.")) {
            year.chop(2);
        }
        year = year.trimmed();

        QLocale russian(QLocale::Russian);
        QString monthName = russian.monthName(monthNum);
        if (!monthName.isEmpty()) {
            monthName[0] = monthName[0].toUpper();
        }

        // ОТЛАДКА
        qDebug() << "=== Файл:" << filepath << "===";
        qDebug() << "E1 raw:" << readCellString(xlsx, 1, 5)
                 << "-> dateVal:" << dateVal
                 << "month:" << monthNum << monthName
                 << "year:" << year;

        // ======== Читаем фамилии H4:H14 ========
        for (int i = 4; i <= 14; i++) {
            QString surname = readCellString(xlsx, i, 8);
            if (surname.isEmpty())
                continue;

            // ОТЛАДКА — что реально лежит в ячейках
            qDebug() << "";
            qDebug() << "--- Строка" << i << "фамилия:" << surname << "---";
            qDebug() << "  F" << i << "(З/п) raw:" << readCellString(xlsx, i, 6)
                     << "-> double:" << readCellDouble(xlsx, i, 6);
            qDebug() << "  I" << i << "(Время) raw:" << readCellString(xlsx, i, 9);

            // Отладка УС
            for (int r = 31; r <= 35; r++) {
                QString n = readCellString(xlsx, r, 1);
                if (!n.isEmpty()) {
                    qDebug() << "  УС: A" << r << "=" << n
                             << " D" << r << "raw:" << readCellString(xlsx, r, 4)
                             << "-> double:" << readCellDouble(xlsx, r, 4);
                }
            }

            // Отладка Аптека
            for (int r = 40; r <= 45; r++) {
                QString n = readCellString(xlsx, r, 1);
                if (!n.isEmpty()) {
                    qDebug() << "  Аптека: A" << r << "=" << n
                             << " B" << r << "raw:" << readCellString(xlsx, r, 2)
                             << "-> double:" << readCellDouble(xlsx, r, 2);
                }
            }

            // Создаём лист если нужно
            if (!surnames.contains(surname)) {
                outputDoc.addSheet(surname);
                surnames.append(surname);
                surnameNextRow[surname] = 3;

                outputDoc.selectSheet(surname);

                outputDoc.mergeCells("B1:F1");
                outputDoc.write("B1", surname, headerFormat);
                outputDoc.write("A1", "", headerFillFormat);
                outputDoc.write("G1", "", headerFillFormat);

                outputDoc.write(2, 1, "Дата",       colHeaderFormat);
                outputDoc.write(2, 2, "З/п Общая",  colHeaderFormat);
                outputDoc.write(2, 3, "УС",         colHeaderFormat);
                outputDoc.write(2, 4, "Аптека",     colHeaderFormat);
                outputDoc.write(2, 5, "Вычет",      colHeaderFormat);
                outputDoc.write(2, 6, "Время",      colHeaderFormat);
                outputDoc.write(2, 7, "Итого",      colHeaderFormat);

                outputDoc.setColumnWidth(1, 18.0);
                outputDoc.setColumnWidth(2, 14.0);
                outputDoc.setColumnWidth(3, 14.0);
                outputDoc.setColumnWidth(4, 14.0);
                outputDoc.setColumnWidth(5, 14.0);
                outputDoc.setColumnWidth(6, 14.0);
                outputDoc.setColumnWidth(7, 14.0);
            }

            outputDoc.selectSheet(surname);
            int row = surnameNextRow[surname];

            // A: Дата
            outputDoc.write(row, 1, monthName + " " + year, dataFormat);

            // B: З/п Общая — столбец F (6)
            double zpObshaya = readCellDouble(xlsx, i, 6);
            outputDoc.write(row, 2, zpObshaya, dataFormat);

            // C: УС — ищем в A31:A35, берём D (4)
            double usVal = 0.0;
            for (int r = 31; r <= 35; r++) {
                QString usName = readCellString(xlsx, r, 1);
                if (!usName.isEmpty() && namesMatch(usName, surname)) {
                    usVal = readCellDouble(xlsx, r, 4);
                    break;
                }
            }
            outputDoc.write(row, 3, usVal, dataFormat);

            // D: Аптека — ищем в A40:A45, берём B (2)
            double aptekaVal = 0.0;
            for (int r = 40; r <= 45; r++) {
                QString apName = readCellString(xlsx, r, 1);
                if (!apName.isEmpty() && namesMatch(apName, surname)) {
                    aptekaVal = readCellDouble(xlsx, r, 2);
                    break;
                }
            }
            outputDoc.write(row, 4, aptekaVal, dataFormat);

            // E: Вычет (пусто)
            outputDoc.write(row, 5, "", dataFormat);

            // F: Время — просто как строку, как есть
            QString vremya = readCellString(xlsx, i, 9);
            outputDoc.write(row, 6, vremya, dataFormat);

            // G: Итого = З/п + УС
            double itogo = zpObshaya + usVal;
            outputDoc.write(row, 7, itogo, dataFormat);

            qDebug() << "  ЗАПИСАНО: З/п=" << zpObshaya
                     << "УС=" << usVal
                     << "Аптека=" << aptekaVal
                     << "Время=" << vremya
                     << "Итого=" << itogo;

            surnameNextRow[surname] = row + 1;
        }
    }

    // ============ Суммы на каждом листе ============
    for (const QString &surname : surnames) {
        outputDoc.selectSheet(surname);
        int sumRow = surnameNextRow[surname];

        outputDoc.write(sumRow, 1, "ИТОГО:", sumFormat);

        double totalZp = 0.0, totalUs = 0.0, totalApteka = 0.0, totalItogo = 0.0;

        for (int r = 3; r < sumRow; r++) {
            auto c2 = outputDoc.cellAt(r, 2);
            if (c2) totalZp += c2->value().toDouble();

            auto c3 = outputDoc.cellAt(r, 3);
            if (c3) totalUs += c3->value().toDouble();

            auto c4 = outputDoc.cellAt(r, 4);
            if (c4) totalApteka += c4->value().toDouble();

            auto c7 = outputDoc.cellAt(r, 7);
            if (c7) totalItogo += c7->value().toDouble();
        }

        outputDoc.write(sumRow, 2, totalZp,     sumFormat);
        outputDoc.write(sumRow, 3, totalUs,      sumFormat);
        outputDoc.write(sumRow, 4, totalApteka,  sumFormat);
        outputDoc.write(sumRow, 5, "",           sumFormat);
        outputDoc.write(sumRow, 6, "",           sumFormat);
        outputDoc.write(sumRow, 7, totalItogo,   sumFormat);
    }

    // ============ Сохранение ============
    if (outputDoc.saveAs(outputFile)) {
        result["success"] = true;
        result["message"] = QString("Обработано файлов: %1\nСоздано листов: %2\nСохранено в:\n%3")
                                .arg(files.size()).arg(surnames.size()).arg(outputFile);
        result["outputPath"] = outputFile;
    } else {
        result["success"] = false;
        result["error"] = "Не удалось сохранить результат в " + outputFile;
    }

    return result;
}
