// EntryIDs are opaque identifiers: equivalent objects may have different
// short/long representations. Never infer equivalence from prefixes/lengths.
internal static class OutlookIdentity
{
    internal static bool AreSame(string? left, string? right, Func<string, string, bool> compare)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        if (string.Equals(left, right, StringComparison.OrdinalIgnoreCase)) return true;
        return compare(left, right); // A provider error must propagate, never authorize a mismatch.
    }
}
