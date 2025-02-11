export interface MessageData {
  messageId: string,
  content:  string,
  neutral: number | null,
  happy: number | null,
  sad: number | null,
  surprise: number | null,
  fear: number | null,
  angry: number | null,
  inappropriate: number | null,
  negative: number | null,
  toxic: number | null,
  informal: number | null,
}