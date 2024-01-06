import Body = GoogleAppsScript.Document.Body;
import Paragraph = GoogleAppsScript.Document.Paragraph;
import ParagraphHeading = GoogleAppsScript.Document.ParagraphHeading;

const onOpen = () => {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();

    const heading1 = updateHeading(body, DocumentApp.ParagraphHeading.HEADING1, null);
    const heading2 = updateHeading(body, DocumentApp.ParagraphHeading.HEADING2, heading1);
    updateHeading(body, DocumentApp.ParagraphHeading.HEADING3, heading2);
};

const HEADING_NUMBER: Record<ParagraphHeading, number> = {
    [DocumentApp.ParagraphHeading.HEADING1]: 1,
    [DocumentApp.ParagraphHeading.HEADING2]: 2,
    [DocumentApp.ParagraphHeading.HEADING3]: 3,
};

const updateHeading = (body: Body, level: ParagraphHeading, parent: Paragraph | null): Paragraph => {
    let heading = findFirstHeading(body, level);
    if (heading && isHeadingUpToDate(heading)) {
        return heading;
    }

    const today = new Date();
    const index = parent ? body.getChildIndex(parent) + 1 : 0;
    const text = formatDate(today, HEADING_NUMBER[level]);
    heading = body.insertParagraph(index, text).setHeading(level);
    Logger.log(`Inserted ${level}: ${text}.`);

    return heading;
};

const findFirstHeading = (body: Body, level: ParagraphHeading): Paragraph | null => {
    const paragraphs = body.getParagraphs();
    for (const paragraph of paragraphs) {
        if (paragraph.getHeading() === level) {
            return paragraph;
        }
    }

    return null;
};

const isHeadingUpToDate = (paragraph: Paragraph): boolean => {
    const date = getDateFromParagraph(paragraph);
    if (date === null) {
        return false;
    }

    const heading = paragraph.getHeading();
    const headingNumber = HEADING_NUMBER[heading];
    if (headingNumber === undefined) {
        return false;
    }

    const dateArray = dateToArray(date);
    const todayArray = dateToArray(new Date());
    for (let i = 0; i < headingNumber; i++) {
        if (todayArray[i] !== dateArray[i]) {
            return false;
        }
    }

    return true;
};

const getDateFromParagraph = (paragraph: Paragraph): Date | null => {
    let date = new Date(paragraph.getText());
    if (date.getTime() > 0) {
        return date;
    }

    const dateElement = paragraph.findElement(DocumentApp.ElementType.DATE).getElement().asDate();
    date = cast<Date>(dateElement.getTimestamp());
    if (date.getTime() > 0) {
        return date;
    }

    return null;
};

const dateToArray = (date: Date): [number, number, number] => {
    return [date.getFullYear(), date.getMonth() + 1, date.getDate()];
};

const formatDate = (date: Date, level: number = 3): string => {
    return date.toLocaleDateString('ja-JP', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).split('/').slice(0, level).join('-');
};

const cast = <T>(x: any): T => x as T;
