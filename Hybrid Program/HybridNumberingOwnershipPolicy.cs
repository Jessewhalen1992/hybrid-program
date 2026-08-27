using System;
using System.Collections.Generic;

namespace HybridSurvey
{
    internal enum LegacyNumberBubbleMatch
    {
        NotForSelectedPolyline,
        AdoptForSelectedPolyline,
        Ambiguous
    }

    internal static class HybridNumberingOwnershipPolicy
    {
        internal static bool CanAdoptResolvedLegacyCandidate(
            bool hasConfirmedOwner,
            int compatibleCandidateCount,
            bool hasAuthoritativeEvidence)
        {
            return compatibleCandidateCount > 0 &&
                (!hasConfirmedOwner || hasAuthoritativeEvidence);
        }

        internal static int[] ReconcileVertexIdsWithLegacyAnchors(
            IList<int> currentIds,
            IList<int> legacyAnchorIds)
        {
            if (currentIds == null) throw new ArgumentNullException(nameof(currentIds));
            if (legacyAnchorIds == null) throw new ArgumentNullException(nameof(legacyAnchorIds));
            if (currentIds.Count != legacyAnchorIds.Count)
                throw new ArgumentException("Current IDs and legacy anchor IDs must have the same length.");

            var claimCounts = new Dictionary<int, int>();
            int maximumId = 0;
            for (int i = 0; i < currentIds.Count; i++)
            {
                maximumId = Math.Max(maximumId, currentIds[i]);
                int anchorId = legacyAnchorIds[i];
                maximumId = Math.Max(maximumId, anchorId);
                if (anchorId <= 0) continue;

                claimCounts.TryGetValue(anchorId, out var count);
                claimCounts[anchorId] = count + 1;
            }

            var resolved = new int[currentIds.Count];
            var used = new HashSet<int>();

            // A unique legacy bubble anchored at a vertex owns that location's
            // persistent identity. Resolve these claims before retaining any
            // newer metadata IDs so a replacement cannot displace the original.
            for (int i = 0; i < legacyAnchorIds.Count; i++)
            {
                int anchorId = legacyAnchorIds[i];
                if (anchorId > 0 && claimCounts[anchorId] == 1 && used.Add(anchorId))
                    resolved[i] = anchorId;
            }

            int nextId = Math.Max(1, maximumId + 1);
            for (int i = 0; i < currentIds.Count; i++)
            {
                if (resolved[i] > 0) continue;

                int currentId = currentIds[i];
                if (currentId > 0 && used.Add(currentId))
                {
                    resolved[i] = currentId;
                    continue;
                }

                while (!used.Add(nextId)) nextId++;
                resolved[i] = nextId++;
            }

            return resolved;
        }

        internal static bool ShouldCreateMissingBubble(
            bool createMissing,
            bool hasKeeper,
            bool insertionSuppressed)
        {
            return createMissing && !hasKeeper && !insertionSuppressed;
        }

        internal static LegacyNumberBubbleMatch ClassifyLegacyBubble(
            bool hasSelectedCandidate,
            bool hasStrongSelectedEvidence,
            bool hasDirectSelectedIdEvidence,
            double selectedDistance,
            bool hasOtherCandidate,
            double closestOtherDistance,
            bool otherMetadataUncertain,
            double tolerance)
        {
            if (!hasSelectedCandidate ||
                double.IsNaN(selectedDistance) ||
                double.IsInfinity(selectedDistance) ||
                selectedDistance < 0.0 ||
                double.IsNaN(tolerance) ||
                double.IsInfinity(tolerance) ||
                tolerance < 0.0)
            {
                return LegacyNumberBubbleMatch.NotForSelectedPolyline;
            }

            // A malformed Hybrid payload on another polyline means ownership
            // cannot be proven. The caller must leave the legacy reference in
            // place and suppress creation rather than guess and duplicate it.
            if (otherMetadataUncertain)
                return LegacyNumberBubbleMatch.Ambiguous;

            if (!hasOtherCandidate)
            {
                // A position match or an otherwise unique persistent ID is
                // conclusive. NUMBER-only evidence can identify which insertion
                // to suppress, but it cannot establish ownership.
                return hasStrongSelectedEvidence || hasDirectSelectedIdEvidence
                    ? LegacyNumberBubbleMatch.AdoptForSelectedPolyline
                    : LegacyNumberBubbleMatch.Ambiguous;
            }

            if (double.IsNaN(closestOtherDistance) ||
                double.IsInfinity(closestOtherDistance) ||
                closestOtherDistance < 0.0)
            {
                return LegacyNumberBubbleMatch.Ambiguous;
            }

            if (selectedDistance + tolerance < closestOtherDistance)
                return LegacyNumberBubbleMatch.AdoptForSelectedPolyline;

            if (closestOtherDistance + tolerance < selectedDistance)
                return LegacyNumberBubbleMatch.NotForSelectedPolyline;

            return LegacyNumberBubbleMatch.Ambiguous;
        }
    }
}
