#include "corrector.h"
#include <QMap>
#include <QStringList>
#include <QCoreApplication>
#include <QEventLoop>

//конструктор по умолчанию
Corrector::Corrector() {}

//функция, которая принимает данные и тип исправляемых данных
Corrector::CorrectionResult Corrector::correct(const QString& text, CorrectionType type)
{
    if (type == EMAIL) {
        return correctEmail(text);
    } else {
        return correctPhone(text);
    }
}

//метод, который прогоняет данные через условия и добавляет в списом
bool Corrector::appendUnique(QStringList& list, const QString& value) const
{
    if(value.isEmpty()) //если пришло пустое значение
    {
        return false;
    }

    if(list.contains(value)) //если значение уже есть в списке
    {
        return false;
    }

    list.append(value); //добавляем значение в список, если прошли все условия
    return true;
};

//исправление почты
Corrector::CorrectionResult Corrector::correctEmail(const QString& text)
{
    CorrectionResult result;
    result.CorrectedText = text;
    result.correctionsCount = 0;

    QString working = text;
    working.replace(QChar(0x00A0), ' '); //неразрывный пробел
    working.replace(QRegularExpression("[\\r\\n\\t]+"), " ");

    //на случай записей вида: info @ site . ru
    working.replace(QRegularExpression(R"(\s*(?:@|\(at\)|\[at\])\s*)", QRegularExpression::CaseInsensitiveOption),"@");

    working.replace(QRegularExpression(R"(\s*(?:\.|\(dot\)|\[dot\])\s*)", QRegularExpression::CaseInsensitiveOption), ".");

    //ищем все email-адреса в тексте
    QRegularExpression emailRegex(R"(([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}))", QRegularExpression::CaseInsensitiveOption);

    QStringList emails;

    QRegularExpressionMatchIterator iterator = emailRegex.globalMatch(working);

    while (iterator.hasNext())
    {
        QRegularExpressionMatch match = iterator.next();

        QString rawEmail = match.captured(1);
        QString normalizedEmail = normalizeEmail(rawEmail, result.changes);

        if (appendUnique(emails, normalizedEmail))
        {
            result.changes << QString("Найден email: \"%1\" -> \"%2\"").arg(rawEmail.trimmed(), normalizedEmail);
        }
    }

    if (emails.isEmpty())
    {
        result.changes << "Email не найден";
        result.CorrectedText = text;
        result.correctionsCount = 0;

        return result;
    }

    //если в ячейке найдено несколько email, записываем каждый через запятую
    result.CorrectedText = emails.join(", ");

    if (result.CorrectedText != text.trimmed())
    {
        result.correctionsCount = 1;
    }

    return result;
}

QString Corrector::normalizeEmail(const QString& rawEmail, QStringList& changes) const
{
    QString email = rawEmail.trimmed().toLower();

    email.remove(QRegularExpression("\\s+"));

    email.remove(QRegularExpression(R"(^[<\(\"'\[]+)"));
    email.remove(QRegularExpression(R"([>\)\"'\],;:]+$)"));

    int atIndex = email.indexOf('@');

    if (atIndex <= 0 || atIndex == email.length() - 1)
    {
        return QString();
    }

    QString localPart = email.left(atIndex);
    QString domain = email.mid(atIndex + 1);

    //исправления частых ошибок в доменах
    QMap<QString, QString> domainFixes;
    domainFixes["gmail.con"] = "gmail.com";
    domainFixes["gamil.com"] = "gmail.com";
    domainFixes["gmial.com"] = "gmail.com";
    domainFixes["gmai.com"] = "gmail.com";

    domainFixes["yandex.ruu"] = "yandex.ru";
    domainFixes["yndex.ru"] = "yandex.ru";
    domainFixes["yandex.ry"] = "yandex.ru";

    domainFixes["mail.ruu"] = "mail.ru";
    domainFixes["bk.ruu"] = "bk.ru";
    domainFixes["list.ruu"] = "list.ru";
    domainFixes["inbox.ruu"] = "inbox.ru";

    if (domainFixes.contains(domain))
    {
        QString oldDomain = domain;
        domain = domainFixes.value(domain);

        changes << QString("Домен исправлен: %1 -> %2").arg(oldDomain, domain);
    }

    if (localPart.isEmpty() || domain.isEmpty() || !domain.contains('.'))
    {
        return QString();
    }

    return localPart + "@" + domain;
}

//исрпавление телефоннных номеров
Corrector::CorrectionResult Corrector::correctPhone(const QString& text)
{
    CorrectionResult result;
    result.CorrectedText = text;
    result.correctionsCount = 0;

    QString working = text;
    working.replace(QChar(0x00A0), ' ');
    working.replace(QRegularExpression("[\\r\\n\\t]+"), " ");

    //ищем номера вида:(812) 441-43-15;+7 (812) 642-86-05;8 (812) 642-86-05;812 642 86 05;+78126428605;88126428605

    QRegularExpression phoneRegex(R"((?:\+?\s*(?:7|8)[\s\-.]*)?(?:\(\s*\d{3,5}\s*\)|\d{3,5})[\s\-.]*\d{2,3}[\s\-.]*\d{2}[\s\-.]*\d{2})");

    QStringList phones;

    QRegularExpressionMatchIterator iterator = phoneRegex.globalMatch(working);

    while (iterator.hasNext())
    {
        QRegularExpressionMatch match = iterator.next();

        QString rawPhone = match.captured(0);
        QString normalizedPhone = normalizePhoneCandidate(rawPhone, result.changes);

        if (appendUnique(phones, normalizedPhone))
        {
            result.changes << QString("Найден телефон: \"%1\" -> \"%2\"").arg(rawPhone.trimmed(), normalizedPhone);
        }
    }

    if (phones.isEmpty())
    {
        result.changes << "Телефон не найден";
        result.CorrectedText = text;
        result.correctionsCount = 0;
        return result;
    }

    //если найдено несколько телефонов, записываем каждый через запятую
    result.CorrectedText = phones.join(", ");

    if (result.CorrectedText != text.trimmed())
    {
        result.correctionsCount = 1;
    }

    return result;
}

QString Corrector::normalizePhoneCandidate(const QString& rawPhone, QStringList& changes) const
{
    QString digits = rawPhone;
    digits.remove(QRegularExpression("[^\\d]"));

    if (digits.isEmpty())
    {
        return QString();
    }

    QString normalized;

    //8 812 441 43 15 -> +78124414315
    if (digits.length() == 11 && digits.startsWith("8"))
    {
        normalized = "+7" + digits.mid(1);
        changes << "Начальная 8 заменена на +7";
    }
    //7 812 441 43 15 -> +78124414315
    else if (digits.length() == 11 && digits.startsWith("7"))
    {
        normalized = "+" + digits;
    }
    //812 441 43 15 -> +78124414315
    else if (digits.length() == 10)
    {
        normalized = "+7" + digits;
        changes << "Добавлен код страны +7";
    }
    else
    {
        //нероссийский формат или мусорный кандидат
        return QString();
    }

    return normalized;
}

