using System.Globalization;

// Pure planning only. Reservations are ADDITIONAL outstanding bytes, not a
// cumulative total already reflected in the physical file. A caller must replace
// a completed reservation with a fresh measured size rather than counting it twice.
internal static class ArchiveVolumePolicy
{
    internal const long KiB = 1024;
    internal const long MiB = 1024 * KiB;
    internal const long GiB = 1024 * MiB;
    internal const long DefaultTargetVolumeBytes = 20 * GiB;
    internal const long DefaultHardLimitBytes = 45 * GiB;
    internal const long OutlookDefaultMaximumBytes = 50 * GiB;
    internal const long DefaultReserveBytes = 512 * MiB;
    internal const long DefaultEmptyVolumeBaseBytes = 271360;
    internal const long MinimumReservationBytes = MiB;

    // An empty PST can grow while its initial folders are created. Pass that
    // measured initial size in emptyVolumeBaseBytes so an oversized first item
    // does not repeatedly request another equally empty file.
    // NewVolume=true is a plan, NOT permission to import without measuring and
    // re-evaluating the newly created file first.
    internal static ArchiveVolumeDecision Decide(
        long currentFileBytes,
        long currentReservedBytes,
        long itemEstimatedBytes,
        long targetVolumeBytes = DefaultTargetVolumeBytes,
        long hardLimitBytes = DefaultHardLimitBytes,
        long reserveBytes = DefaultReserveBytes,
        long emptyVolumeBaseBytes = DefaultEmptyVolumeBaseBytes)
    {
        if (currentFileBytes < 0) return Deny("invalid-current-file-size");
        if (currentReservedBytes < 0) return Deny("invalid-outstanding-reservation-size");
        if (itemEstimatedBytes < 0) return Deny("invalid-item-size");
        if (targetVolumeBytes <= 0 || hardLimitBytes <= 0 || reserveBytes < 0 || emptyVolumeBaseBytes < 0)
            return Deny("invalid-volume-policy");
        if (hardLimitBytes > OutlookDefaultMaximumBytes) return Deny("hard-limit-exceeds-outlook-default");
        if (reserveBytes >= hardLimitBytes) return Deny("reserve-consumes-hard-limit");

        long usableLimit = hardLimitBytes - reserveBytes;
        if (targetVolumeBytes > usableLimit) return Deny("target-exceeds-usable-hard-limit");
        if (emptyVolumeBaseBytes >= usableLimit) return Deny("empty-volume-consumes-usable-hard-limit");

        long reservation;
        try
        {
            // ceil(n * 1.25), with integer arithmetic and checked overflow.
            long quarter = itemEstimatedBytes / 4 + (itemEstimatedBytes % 4 == 0 ? 0 : 1);
            reservation = Math.Max(MinimumReservationBytes, checked(itemEstimatedBytes + quarter));
        }
        catch (OverflowException) { return Deny("item-reservation-overflow"); }

        bool empty = currentReservedBytes == 0 && currentFileBytes <= emptyVolumeBaseBytes;
        // Zero can represent a not-yet-created PST: do not ignore its base size.
        long effectiveFileBytes = empty ? Math.Max(currentFileBytes, emptyVolumeBaseBytes) : currentFileBytes;
        long currentBudget, projected, freshProjected;
        try
        {
            currentBudget = checked(effectiveFileBytes + currentReservedBytes);
            projected = checked(currentBudget + reservation);
            freshProjected = checked(emptyVolumeBaseBytes + reservation);
        }
        catch (OverflowException) { return Deny("volume-budget-overflow", reservation); }

        // Existing observations or reservations above the bound are an invariant
        // failure to investigate, not a reason to silently continue in a new PST.
        if (currentBudget > usableLimit) return Deny("existing-volume-budget-exceeds-hard-limit", reservation);
        if (freshProjected > usableLimit) return Deny("single-item-exceeds-empty-volume-hard-limit", reservation);

        if (projected <= targetVolumeBytes)
            return new(false, reservation, true, "fits-current-volume");

        if (empty)
            return new(false, reservation, true, "dedicated-oversized-item-in-empty-volume");

        return new(true, reservation, true,
            projected > usableLimit ? "roll-before-hard-limit" :
            freshProjected > targetVolumeBytes ? "roll-for-dedicated-oversized-item" : "roll-before-target-limit");
    }

    internal static string GetFileName(Guid jobId, int partNumber)
    {
        if (jobId == Guid.Empty) throw new ArgumentException("A execução deve possuir um GUID válido.", nameof(jobId));
        if (partNumber <= 0) throw new ArgumentOutOfRangeException(nameof(partNumber), "O número da parte deve ser positivo.");
        return "archive-" + jobId.ToString("N") + "-part-" + partNumber.ToString("D4", CultureInfo.InvariantCulture) + ".pst";
    }

    internal static string GetFileName(string jobId, int partNumber)
    {
        if (!Guid.TryParse(jobId, out var parsed)) throw new ArgumentException("A execução deve possuir um GUID válido.", nameof(jobId));
        return GetFileName(parsed, partNumber);
    }

    private static ArchiveVolumeDecision Deny(string reason, long reservation = 0) => new(false, reservation, false, reason);

    internal static void SelfTest()
    {
        static void Check(bool condition, string message)
        { if (!condition) throw new Exception("Política de volumes: " + message); }

        var normal = Decide(DefaultEmptyVolumeBaseBytes, 0, 100);
        Check(normal.Allowed && !normal.NewVolume && normal.ReservationBytes == MiB, "reserva mínima ou volume inicial incorreto.");
        var rounded = Decide(DefaultEmptyVolumeBaseBytes, 0, 4 * MiB + 1);
        Check(rounded.ReservationBytes == 5 * MiB + 2, "estimativa de 125% não arredondada para cima.");

        long target = 10 * MiB, hard = 100 * MiB, reserve = MiB;
        var exactTarget = Decide(DefaultEmptyVolumeBaseBytes, 0, 0,
            DefaultEmptyVolumeBaseBytes + MiB, hard, reserve);
        Check(exactTarget.Allowed && !exactTarget.NewVolume, "limite exato da meta rejeitado.");
        var roll = Decide(target - MiB, 0, MiB, target, hard, reserve);
        Check(roll.Allowed && roll.NewVolume, "não houve rotação antes de ultrapassar a meta.");
        var reserved = Decide(6 * MiB, MiB, MiB, 8 * MiB, hard, reserve);
        Check(reserved.Allowed && reserved.NewVolume, "reserva pendente foi substituída pelo tamanho físico em vez de somada.");

        var dedicated = Decide(DefaultEmptyVolumeBaseBytes, 0, 6 * MiB, 2 * MiB, 10 * MiB, MiB);
        Check(dedicated.Allowed && !dedicated.NewVolume && dedicated.ReservationBytes > 2 * MiB,
            "item acima da meta não ficou em volume vazio dedicado.");
        var oversized = Decide(DefaultEmptyVolumeBaseBytes, 0, 8 * MiB, 2 * MiB, 10 * MiB, MiB);
        Check(!oversized.Allowed && !oversized.NewVolume, "item acima do limite seguro foi aceito.");
        var freshHard = Decide(DefaultEmptyVolumeBaseBytes, 0, 0, MiB,
            DefaultEmptyVolumeBaseBytes + MiB + reserve, reserve);
        Check(freshHard.Allowed && !freshHard.NewVolume, "limite seguro exato foi rejeitado.");
        var measuredEmpty = Decide(2 * MiB, 0, 0, MiB, 32 * MiB, MiB, 2 * MiB);
        Check(measuredEmpty.Allowed && !measuredEmpty.NewVolume, "PST vazio medido entrou em rotação repetida.");
        var fullCurrent = Decide(9 * MiB, 0, 0, 2 * MiB, 10 * MiB, MiB);
        Check(fullCurrent.Allowed && fullCurrent.NewVolume, "volume cheio dentro do limite não iniciou nova parte.");
        var alreadyExceeded = Decide(9 * MiB + 1, 0, 0, 2 * MiB, 10 * MiB, MiB);
        Check(!alreadyExceeded.Allowed && !alreadyExceeded.NewVolume, "tamanho real já acima do limite foi ocultado por rotação.");
        Check(!Decide(-1, 0, 1).Allowed && !Decide(0, -1, 1).Allowed && !Decide(0, 0, -1).Allowed,
            "entrada negativa aceita.");
        Check(!Decide(0, 0, 1, 0).Allowed && !Decide(0, 0, 1, hard, hard, 1).Allowed &&
            !Decide(0, 0, 1, MiB, MiB, MiB).Allowed &&
            !Decide(0, 0, 1, GiB, OutlookDefaultMaximumBytes + 1, MiB).Allowed, "configuração inválida aceita.");
        Check(!Decide(0, 0, long.MaxValue).Allowed && !Decide(long.MaxValue, long.MaxValue, 0).Allowed,
            "overflow gerou permissão falsa.");

        var id = Guid.Parse("d19693ce-86f1-42a9-b419-a549debd91af");
        Check(GetFileName(id, 1) == "archive-d19693ce86f142a9b419a549debd91af-part-0001.pst", "nome de parte não determinístico.");
        Check(GetFileName(id.ToString(), 10000).EndsWith("-part-10000.pst", StringComparison.Ordinal), "número de parte foi truncado.");
        try { GetFileName("../escape", 1); throw new Exception("Identificador inválido aceito."); }
        catch (ArgumentException) { }
        try { GetFileName(id, 0); throw new Exception("Parte zero aceita."); }
        catch (ArgumentOutOfRangeException) { }
        Console.WriteLine("Volumes: metas/limites, reserva pendente, item dedicado, overflow e nome seguro conferidos. Sem Outlook ou arquivos.");
    }
}

internal sealed record ArchiveVolumeDecision(bool NewVolume, long ReservationBytes, bool Allowed, string Reason);
