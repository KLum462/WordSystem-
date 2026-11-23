using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using WordClient.Core.Models;

namespace WordServer.Api.Services
{
    public class TextProcessingService
    {
        public static readonly string[] SupportedOperations =
        {
            "uppercase",
            "lowercase",
            "trim",
            "reverse",
            "word_count",
            "title_case",
            "capitalize",
            "normalize",
            "anonymize",
            "extract_entities",
            "summarize",
            "stats"
        };

        private static readonly Regex EmailRegex = new(
            @"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
            RegexOptions.Compiled);

        private static readonly Regex PhoneRegex = new(
            @"(?<!\d)(?:\+?\d{1,3}[\s\-]?)?(?:\(?\d{3}\)?[\s\-]?)\d{3}[\s\-]?\d{2}[\s\-]?\d{2}(?!\d)",
            RegexOptions.Compiled);

        private static readonly Regex CardRegex = new(
            @"(?<!\d)\d{4}[\s\-]?\d{4}[\s\-]?\d{4}[\s\-]?\d{4}(?!\d)",
            RegexOptions.Compiled);

        private static readonly Regex InnRegex = new(
            @"(?i)(ИНН|INN)\s*[:№#]?\s*\d{10,12}",
            RegexOptions.Compiled);

        private static readonly Regex PassportRegex = new(
            @"(?i)(паспорт|passport)\s*[:№#]?\s*\d{2}\s?\d{2}\s+\d{6}",
            RegexOptions.Compiled);

        private static readonly Regex DateRegex = new(
            @"\b(?:\d{2}\.\d{2}\.\d{4}|\d{4}-\d{2}-\d{2})\b",
            RegexOptions.Compiled);

        private static readonly Regex MoneyRegex = new(
            @"(?i)(?:\b\d{1,3}(?:[\s\u00A0]\d{3})*(?:[.,]\d{1,2})?\s?(?:руб\.?|₽|RUB|USD|EUR|\$|€)|(?:\$|€)\s?\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?)",
            RegexOptions.Compiled);

        private static readonly Regex UrlRegex = new(
            @"https?://[^\s<>""'\)\]]+",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
        {
            "the", "and", "for", "with", "that", "this", "from", "are", "was", "were", "have", "has", "not", "but", "into", "about", "your", "our", "will",
            "и", "в", "во", "не", "что", "он", "на", "я", "с", "со", "как", "а", "то", "все", "она", "так", "его", "но", "да", "ты", "к", "у", "же",
            "вы", "за", "бы", "по", "только", "ее", "мне", "было", "вот", "от", "меня", "еще", "нет", "о", "из", "ему", "теперь", "когда", "даже",
            "ну", "вдруг", "ли", "если", "уже", "или", "ни", "быть", "был", "него", "до", "вас", "нибудь", "опять", "уж", "вам", "ведь", "там", "потом",
            "себя", "ничего", "ей", "может", "они", "тут", "где", "есть", "надо", "ней", "для", "мы", "тебя", "их", "чем", "была", "сам", "чтоб", "без",
            "будто", "чего", "раз", "тоже", "себе", "под", "будет", "ж", "тогда", "кто", "этот", "того", "потому", "этого", "какой", "совсем", "ним",
            "здесь", "этом", "один", "почти", "мой", "тем", "чтобы", "нее", "сейчас", "были", "куда", "зачем", "всех", "никогда", "можно", "при",
            "наконец", "два", "об", "другой", "хоть", "после", "над", "больше", "тот", "через", "эти", "нас", "про", "всего", "них", "какая", "много"
        };

        public bool TryProcess(string text, string? rawOperation, out string result, out string error)
        {
            var operation = rawOperation?.Trim().ToLowerInvariant() ?? "uppercase";
            error = string.Empty;

            switch (operation)
            {
                case "uppercase":
                    result = text.ToUpperInvariant();
                    return true;

                case "lowercase":
                    result = text.ToLowerInvariant();
                    return true;

                case "trim":
                    result = text.Trim();
                    return true;

                case "reverse":
                    var chars = text.ToCharArray();
                    Array.Reverse(chars);
                    result = new string(chars);
                    return true;

                case "word_count":
                    var words = SplitWords(text);
                    result = $"Word count: {words.Length}";
                    return true;

                case "title_case":
                    result = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLowerInvariant());
                    return true;

                case "capitalize":
                    result = text.Length > 0
                        ? char.ToUpperInvariant(text[0]) + text.Substring(1)
                        : string.Empty;
                    return true;

                case "normalize":
                    result = NormalizeText(text);
                    return true;

                case "anonymize":
                    result = AnonymizePii(text);
                    return true;

                case "extract_entities":
                    result = ExtractEntities(text);
                    return true;

                case "summarize":
                    result = SummarizeText(text);
                    return true;

                case "stats":
                    result = ComputeDetailedStats(text);
                    return true;

                default:
                    result = string.Empty;
                    error = $"Unsupported operation '{rawOperation}'. Supported operations: {string.Join(", ", SupportedOperations)}.";
                    return false;
            }
        }

        public DocumentCheckResult InspectDocument(DocumentCheckRequest request)
        {
            var text = request.Text ?? string.Empty;
            var paragraphsList = request.Paragraphs
                ?? text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).ToList();
            var words = SplitWords(text);

            int emptyParagraphs = paragraphsList.Count(p => string.IsNullOrWhiteSpace(p));
            var issues = new List<string>();

            if (emptyParagraphs > 0)
            {
                issues.Add($"Found {emptyParagraphs} empty paragraph(s) in document.");
            }

            if (words.Length == 0)
            {
                issues.Add("Document contains no readable words.");
            }

            int doubleSpaceCount = Regex.Matches(text, @"[^\r\n] {2,}[^\r\n]").Count;
            if (doubleSpaceCount > 0)
            {
                issues.Add($"Found {doubleSpaceCount} occurrence(s) of multiple consecutive spaces.");
            }

            int trailingSpaceParagraphs = paragraphsList.Count(p => !string.IsNullOrWhiteSpace(p) && (p.EndsWith(" ") || p.EndsWith("\t")));
            if (trailingSpaceParagraphs > 0)
            {
                issues.Add($"Found {trailingSpaceParagraphs} paragraph(s) with trailing whitespace.");
            }

            var sentences = SplitSentences(text);
            int longSentences = sentences.Count(s => SplitWords(s).Length > 40);
            if (longSentences > 0)
            {
                issues.Add($"Found {longSentences} overly long sentence(s) (>40 words) — consider splitting for readability.");
            }

            int emailCount = EmailRegex.Matches(text).Count;
            int phoneCount = PhoneRegex.Matches(text).Count;
            if (emailCount > 0 || phoneCount > 0)
            {
                issues.Add($"Potential PII detected ({emailCount} email(s), {phoneCount} phone(s)) — consider running 'anonymize' before external sharing.");
            }

            return new DocumentCheckResult
            {
                WordCount = words.Length,
                ParagraphCount = paragraphsList.Count,
                EmptyParagraphCount = emptyParagraphs,
                TableCount = request.TableCount ?? 0,
                CharacterCount = text.Length,
                Issues = issues,
                Summary = $"Document inspection complete: {words.Length} words, {paragraphsList.Count} paragraphs ({emptyParagraphs} empty), {request.TableCount ?? 0} tables, {issues.Count} issue(s)."
            };
        }

        private static string NormalizeText(string text)
        {
            var lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var cleanedLines = new List<string>();
            bool previousWasEmpty = false;

            foreach (var line in lines)
            {
                var collapsed = Regex.Replace(line.Trim(), @"[ \t]{2,}", " ");
                collapsed = Regex.Replace(collapsed, @"\s+([,.;:!?])", "$1");

                if (string.IsNullOrWhiteSpace(collapsed))
                {
                    if (!previousWasEmpty && cleanedLines.Count > 0)
                    {
                        cleanedLines.Add(string.Empty);
                        previousWasEmpty = true;
                    }
                }
                else
                {
                    cleanedLines.Add(collapsed);
                    previousWasEmpty = false;
                }
            }

            while (cleanedLines.Count > 0 && string.IsNullOrEmpty(cleanedLines[^1]))
            {
                cleanedLines.RemoveAt(cleanedLines.Count - 1);
            }

            return string.Join(Environment.NewLine, cleanedLines);
        }

        private static string AnonymizePii(string text)
        {
            var result = EmailRegex.Replace(text, "[EMAIL]");
            result = CardRegex.Replace(result, "[CARD]");
            result = InnRegex.Replace(result, "$1: [INN]");
            result = PassportRegex.Replace(result, "$1: [PASSPORT]");
            result = PhoneRegex.Replace(result, "[PHONE]");
            return result;
        }

        private static string ExtractEntities(string text)
        {
            var emails = EmailRegex.Matches(text).Select(m => m.Value).Distinct().ToList();
            var phones = PhoneRegex.Matches(text).Select(m => m.Value.Trim()).Distinct().ToList();
            var dates = DateRegex.Matches(text).Select(m => m.Value).Distinct().ToList();
            var amounts = MoneyRegex.Matches(text).Select(m => m.Value.Trim()).Distinct().ToList();
            var urls = UrlRegex.Matches(text).Select(m => m.Value).Distinct().ToList();

            var sb = new StringBuilder();
            sb.AppendLine("=== Extracted Document Entities ===");
            sb.AppendLine($"Emails ({emails.Count}): {(emails.Count > 0 ? string.Join(", ", emails) : "none")}");
            sb.AppendLine($"Phones ({phones.Count}): {(phones.Count > 0 ? string.Join(", ", phones) : "none")}");
            sb.AppendLine($"Dates ({dates.Count}): {(dates.Count > 0 ? string.Join(", ", dates) : "none")}");
            sb.AppendLine($"Amounts ({amounts.Count}): {(amounts.Count > 0 ? string.Join(", ", amounts) : "none")}");
            sb.Append($"URLs ({urls.Count}): {(urls.Count > 0 ? string.Join(", ", urls) : "none")}");
            return sb.ToString();
        }

        private static string SummarizeText(string text)
        {
            var sentences = SplitSentences(text);
            if (sentences.Count <= 2)
            {
                return text.Trim();
            }

            var frequencies = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var word in SplitWords(text))
            {
                var clean = CleanWord(word);
                if (clean.Length >= 3 && !StopWords.Contains(clean))
                {
                    frequencies[clean] = frequencies.TryGetValue(clean, out var count) ? count + 1 : 1;
                }
            }

            int takeCount = Math.Min(3, Math.Max(1, (int)Math.Ceiling(sentences.Count * 0.35)));
            var ranked = sentences
                .Select((sentence, index) => new
                {
                    Index = index,
                    Sentence = sentence,
                    Score = ScoreSentence(sentence, frequencies, index)
                })
                .OrderByDescending(x => x.Score)
                .Take(takeCount)
                .OrderBy(x => x.Index)
                .Select(x => x.Sentence)
                .ToList();

            return string.Join(" ", ranked);
        }

        private static string ComputeDetailedStats(string text)
        {
            var words = SplitWords(text);
            var sentences = SplitSentences(text);
            var cleanedWords = words
                .Select(CleanWord)
                .Where(w => w.Length >= 3 && !StopWords.Contains(w))
                .ToList();

            int uniqueWords = words.Select(CleanWord).Where(w => w.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            double avgWordsPerSentence = sentences.Count > 0 ? (double)words.Length / sentences.Count : words.Length;
            int readingTimeSeconds = Math.Max(1, (int)Math.Ceiling(words.Length / 200.0 * 60.0));

            var topWords = cleanedWords
                .GroupBy(w => w.ToLowerInvariant())
                .OrderByDescending(g => g.Count())
                .Take(5)
                .Select(g => $"{g.Key} ({g.Count()})")
                .ToList();

            var sb = new StringBuilder();
            sb.AppendLine("=== Text Analytics ===");
            sb.AppendLine($"Words: {words.Length} (unique: {uniqueWords})");
            sb.AppendLine($"Sentences: {sentences.Count} (avg {avgWordsPerSentence:F1} words/sentence)");
            sb.AppendLine($"Characters: {text.Length}");
            sb.AppendLine($"Estimated reading time: ~{readingTimeSeconds} sec");
            sb.Append($"Top keywords: {(topWords.Count > 0 ? string.Join(", ", topWords) : "n/a")}");
            return sb.ToString();
        }

        private static double ScoreSentence(string sentence, Dictionary<string, int> frequencies, int index)
        {
            var words = SplitWords(sentence);
            if (words.Length == 0) return 0;

            double score = 0;
            foreach (var word in words)
            {
                var clean = CleanWord(word);
                if (frequencies.TryGetValue(clean, out var freq))
                {
                    score += freq;
                }
            }

            // Slight preference for opening sentence and moderate normalization by length
            double normalized = score / Math.Sqrt(words.Length);
            if (index == 0) normalized *= 1.15;
            return normalized;
        }

        private static string[] SplitWords(string text)
        {
            return text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static List<string> SplitSentences(string text)
        {
            return Regex.Split(text.Trim(), @"(?<=[.!?])\s+|\r?\n+")
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToList();
        }

        private static string CleanWord(string word)
        {
            return Regex.Replace(word, @"[^\p{L}\p{Nd}]", string.Empty);
        }
    }
}
