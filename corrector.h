#ifndef CORRECTOR_H
#define CORRECTOR_H

#include <QString>
#include <QRegularExpression> //для контроля ввода текста в поля LineEdit
#include <QStringList>

class Corrector
{
public:
    Corrector();

    enum CorrectionType {EMAIL, PHONE};

    struct CorrectionResult
    {
        QString CorrectedText;
        int correctionsCount;
        QStringList changes;
    };

    Corrector::CorrectionResult correct(const QString& text, CorrectionType type);

private:
    CorrectionResult correctEmail(const QString& email);
    CorrectionResult correctPhone(const QString& phone);

    QString normalizeEmail(const QString& rawEmail, QStringList& changes) const;
    QString normalizePhoneCandidate(const QString& rawPhone, QStringList& changes) const;

    bool appendUnique(QStringList& list, const QString& value) const;
};

#endif // CORRECTOR_H
