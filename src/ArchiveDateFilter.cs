// MAPI PT_SYSTIME is UTC even when COM returns DateTimeKind.Unspecified.
// Compare calendar dates in the persisted zone: no fragile 23:59:59 bound,
// no loss of sub-second precision, and no assumption that every day has 24 h.
internal static class ArchiveDateFilter
{
    internal static (DateTime? Utc, string Basis) SelectDate(DateTime? receivedUtc, DateTime? sentUtc)
    {
        // FROMME also occurs in received mail sent to oneself. A real delivery
        // timestamp takes precedence; sent-only items fall back to submission.
        bool Valid(DateTime? date) => date is { Year: >= 1601 and < 4500 } value && value.Kind != DateTimeKind.Local;
        return Valid(receivedUtc) ? (receivedUtc, "received") : Valid(sentUtc) ? (sentUtc, "sent") : (null, "unavailable");
    }
    internal static bool Includes(DateTime? messageDateUtc, DateOnly? cutoff, string timeZoneId)
    {
        if (cutoff is null) return true;
        if (messageDateUtc is not { } value || value.Year is < 1601 or >= 4500 || value.Kind == DateTimeKind.Local)
            throw new ArchiveReviewException("Mensagem sem data UTC válida para aplicar o corte; nenhum plano parcial será criado.");
        var zone = TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);
        var local = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(value, DateTimeKind.Utc), zone);
        return DateOnly.FromDateTime(local) <= cutoff.Value;
    }
}
