// The adapter's own words on a question card (labels, errors, confirmations).
// Questions, options and results are the engine's and are never rewritten.

export interface QuestionCardTexts {
  submit: string;
  submitIdea: string;
  change: string;
  addAnother: string;
  volunteer: string;
  proposePerson: string;
  yourAnswer: string;
  yourContributions(count: number): string;
  answeredAt(time: string): string;
  changeableUntilClose: string;
  openUntil(time: string): string;
  closedAt(time: string): string;
  noLongerActive: string;
  proposalChanged: string;
  notForThisConversation: string;
  unverifiedAnswerer: string;
  couldNotRead: string;
  required: string;
  chooseOne: string;
  unknownOption: string;
  rankLabel(place: number): string;
  rankingIncomplete: string;
  rankingDuplicate(label: string, places: readonly number[]): string;
  tooFew(min: number): string;
  tooMany(max: number): string;
  agree: string;
  object: string;
  reasonLabel: string;
  reasonRequired: string;
  textTooLong(max: number): string;
  pickSomeone: string;
  unknownPerson: string;
  self: string;
  fixErrors: string;
}

const DE: QuestionCardTexts = {
  submit: 'Antwort senden',
  submitIdea: 'Beitrag senden',
  change: 'Ändern',
  addAnother: 'Weiteren Beitrag senden',
  volunteer: 'Ich übernehme das',
  proposePerson: 'Vorschlag senden',
  yourAnswer: 'Deine Antwort',
  yourContributions: (count) => `Deine Beiträge (${String(count)})`,
  answeredAt: (time) => `Gesendet ${time}`,
  changeableUntilClose: 'Deine Antwort kannst du bis zum Ende ändern.',
  openUntil: (time) => `Offen bis ${time}.`,
  closedAt: (time) => `Geschlossen am ${time}.`,
  noLongerActive: 'Diese Abstimmung ist nicht mehr aktiv.',
  proposalChanged: 'Der Vorschlag wurde inzwischen geändert. Bitte antworte auf die aktuelle Fassung.',
  notForThisConversation: 'Diese Karte gehört nicht in diese Unterhaltung. Deine Antwort wurde nicht übernommen.',
  unverifiedAnswerer:
    'Deine Antwort wurde übermittelt. Teams hat nicht mitgeteilt, wer du bist, deshalb kann sie dir nicht sicher zugeordnet werden.',
  couldNotRead: 'Diese Antwort konnte nicht gelesen werden. Bitte öffne die Karte erneut.',
  required: 'Bitte ausfüllen.',
  chooseOne: 'Bitte eine Option wählen.',
  unknownOption: 'Diese Option gibt es auf dieser Karte nicht.',
  rankLabel: (place) => `${String(place)}. Platz`,
  rankingIncomplete: 'Bitte jeden Platz vergeben.',
  rankingDuplicate: (label, places) =>
    `Jede Option darf nur einmal vorkommen: „${label}“ steht auf Platz ${places.map(String).join(' und ')}.`,
  tooFew: (min) => `Bitte mindestens ${String(min)} auswählen.`,
  tooMany: (max) => `Bitte höchstens ${String(max)} auswählen.`,
  agree: 'Einverstanden',
  object: 'Einwand',
  reasonLabel: 'Begründung (bei Einwand nötig)',
  reasonRequired: 'Ein Einwand braucht eine Begründung.',
  textTooLong: (max) => `Bitte höchstens ${String(max)} Zeichen.`,
  pickSomeone: 'Bitte eine Person wählen oder selbst übernehmen.',
  unknownPerson: 'Diese Person konnte nicht zugeordnet werden.',
  self: 'ich selbst',
  fixErrors: 'Bitte prüfe die markierten Angaben.',
};

const EN: QuestionCardTexts = {
  submit: 'Send answer',
  submitIdea: 'Send contribution',
  change: 'Change',
  addAnother: 'Send another contribution',
  volunteer: 'I will take this',
  proposePerson: 'Send proposal',
  yourAnswer: 'Your answer',
  yourContributions: (count) => `Your contributions (${String(count)})`,
  answeredAt: (time) => `Sent ${time}`,
  changeableUntilClose: 'You can change your answer until the close.',
  openUntil: (time) => `Open until ${time}.`,
  closedAt: (time) => `Closed on ${time}.`,
  noLongerActive: 'This vote is no longer active.',
  proposalChanged: 'The proposal has changed since. Please answer the current version.',
  notForThisConversation: 'This card does not belong to this conversation. Your answer was not taken.',
  unverifiedAnswerer:
    'Your answer was passed on. Teams did not state who you are, so it cannot be attributed to you with certainty.',
  couldNotRead: 'This answer could not be read. Please open the card again.',
  required: 'Please fill this in.',
  chooseOne: 'Please choose one option.',
  unknownOption: 'This option is not on this card.',
  rankLabel: (place) => `Place ${String(place)}`,
  rankingIncomplete: 'Please fill every place.',
  rankingDuplicate: (label, places) =>
    `Each option may appear only once: "${label}" is on place ${places.map(String).join(' and ')}.`,
  tooFew: (min) => `Please choose at least ${String(min)}.`,
  tooMany: (max) => `Please choose at most ${String(max)}.`,
  agree: 'Agree',
  object: 'Object',
  reasonLabel: 'Reason (needed for an objection)',
  reasonRequired: 'An objection needs a reason.',
  textTooLong: (max) => `Please use at most ${String(max)} characters.`,
  pickSomeone: 'Please choose a person or take it yourself.',
  unknownPerson: 'This person could not be identified.',
  self: 'myself',
  fixErrors: 'Please check the marked fields.',
};

/** German unless the card states another language: the adapter's other
 *  texts are German, and an instance's users read one language. */
export function questionCardTexts(locale: string | undefined): QuestionCardTexts {
  return locale !== undefined && locale.toLowerCase().startsWith('en') ? EN : DE;
}

/** A time for a card line, in the card's language. */
export function formatCardTime(iso: string, locale: string | undefined): string {
  const time = new Date(iso);
  if (Number.isNaN(time.getTime())) return iso;
  try {
    return new Intl.DateTimeFormat(locale?.toLowerCase().startsWith('en') ? 'en-GB' : 'de-DE', {
      dateStyle: 'medium',
      timeStyle: 'short',
      timeZone: 'Europe/Berlin',
    }).format(time);
  } catch {
    return iso;
  }
}
