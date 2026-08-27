using HybridSurvey;
using System;
using System.Globalization;
using System.IO;

namespace HybridProgram.DecisionTests
{
    internal static class Program
    {
        private static int Main()
        {
            try
            {
                TestCoordinateStorageAndDisplay();
                TestCoordinateEditPreservesHiddenPrecision();
                TestInvariantCoordinateParsing();
                TestDuplicateCoordinateIdentity();
                TestMarkerInsertionPolicy();
                TestLegacyNumberBubbleOwnershipPolicy();
                TestLegacyNumberBubbleOrchestrationPolicy();
                TestLegacyNumberBubbleIdentityReconciliation();
                TestJtbUtmDxfScaleReader();
                Console.WriteLine("HybridProgram decision tests passed.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return 1;
            }
        }

        private static void TestCoordinateStorageAndDisplay()
        {
            AssertNear(123.457, HybridValuePolicy.StoreCoordinate(123.4565), nameof(TestCoordinateStorageAndDisplay));
            AssertNear(-123.457, HybridValuePolicy.StoreCoordinate(-123.4565), nameof(TestCoordinateStorageAndDisplay));

            string display = HybridValuePolicy.DisplayCoordinate(123.456);
            if (!string.Equals(display, "123.46", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    nameof(TestCoordinateStorageAndDisplay) + ": expected F2 display, got " + display + ".");
        }

        private static void TestInvariantCoordinateParsing()
        {
            if (!HybridValuePolicy.TryParseCoordinate("123.4567", out var coordinate))
                throw new InvalidOperationException(nameof(TestInvariantCoordinateParsing) + ": dot decimal was rejected.");
            AssertNear(123.457, coordinate, nameof(TestInvariantCoordinateParsing));

            if (HybridValuePolicy.TryParseCoordinate("123,456", out _))
                throw new InvalidOperationException(nameof(TestInvariantCoordinateParsing) + ": comma decimal was accepted.");
            if (HybridValuePolicy.TryParseCoordinate("NaN", out _))
                throw new InvalidOperationException(nameof(TestInvariantCoordinateParsing) + ": NaN was accepted.");
        }

        private static void TestDuplicateCoordinateIdentity()
        {
            string first = HybridValuePolicy.StoredCoordinateKey(326046.3204, 6160960.4304);
            string duplicate = HybridValuePolicy.StoredCoordinateKey(326046.32049, 6160960.43049);
            string distinct = HybridValuePolicy.StoredCoordinateKey(326046.321, 6160960.430);

            if (!string.Equals(first, duplicate, StringComparison.Ordinal) ||
                string.Equals(first, distinct, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    nameof(TestDuplicateCoordinateIdentity) +
                    ": duplicate X/Y identity did not follow 0.001 storage precision.");
            }
        }

        private static void TestCoordinateEditPreservesHiddenPrecision()
        {
            if (HybridValuePolicy.StorageCoordinateText(123.456) != "123.456")
                throw new InvalidOperationException(
                    nameof(TestCoordinateEditPreservesHiddenPrecision) + ": clipboard storage text lost precision.");

            if (!HybridValuePolicy.TryResolveEditedCoordinate(
                    "123.46",
                    123.456,
                    "123.46",
                    out var unchanged))
                throw new InvalidOperationException(
                    nameof(TestCoordinateEditPreservesHiddenPrecision) + ": unchanged edit was rejected.");
            AssertNear(123.456, unchanged, nameof(TestCoordinateEditPreservesHiddenPrecision));

            if (!HybridValuePolicy.TryResolveEditedCoordinate(
                    "123.47",
                    123.456,
                    "123.46",
                    out var changed))
                throw new InvalidOperationException(
                    nameof(TestCoordinateEditPreservesHiddenPrecision) + ": changed edit was rejected.");
            AssertNear(123.470, changed, nameof(TestCoordinateEditPreservesHiddenPrecision));
        }

        private static void TestMarkerInsertionPolicy()
        {
            if (HybridMarkerPolicy.DesiredMarkerType("XC", true) != "XC" ||
                HybridMarkerPolicy.DesiredMarkerType("ec", true) != "EC")
                throw new InvalidOperationException(nameof(TestMarkerInsertionPolicy) + ": XC/EC were not retained.");

            if (HybridMarkerPolicy.DesiredMarkerType("RC", true).Length != 0)
                throw new InvalidOperationException(nameof(TestMarkerInsertionPolicy) + ": RC would receive a marker block.");
            if (HybridMarkerPolicy.DesiredMarkerType("XC", false).Length != 0)
                throw new InvalidOperationException(nameof(TestMarkerInsertionPolicy) + ": disabled marker updates would insert a block.");
        }

        private static void TestLegacyNumberBubbleOwnershipPolicy()
        {
            const double tolerance = 0.03;

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.AdoptForSelectedPolyline,
                true,
                false,
                true,
                12.0,
                false,
                double.PositiveInfinity,
                false,
                tolerance,
                "moved bubble has a unique persistent ID");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.AdoptForSelectedPolyline,
                true,
                false,
                false,
                5.0,
                true,
                25.0,
                false,
                tolerance,
                "moved bubble is clearly closest to selected polyline");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.NotForSelectedPolyline,
                true,
                true,
                true,
                25.0,
                true,
                5.0,
                false,
                tolerance,
                "moved bubble is clearly closest to another polyline");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.Ambiguous,
                true,
                true,
                true,
                10.0,
                true,
                10.01,
                false,
                tolerance,
                "near-tie suppresses replacement insertion");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.Ambiguous,
                true,
                true,
                true,
                0.0,
                false,
                double.PositiveInfinity,
                true,
                tolerance,
                "unreadable competing metadata suppresses replacement insertion");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.NotForSelectedPolyline,
                false,
                false,
                false,
                double.PositiveInfinity,
                false,
                double.PositiveInfinity,
                false,
                tolerance,
                "unmatched legacy bubble");

            AssertLegacyMatch(
                LegacyNumberBubbleMatch.Ambiguous,
                true,
                false,
                false,
                50.0,
                false,
                double.PositiveInfinity,
                false,
                tolerance,
                "NUMBER-only moved bubble suppresses insertion without being claimed");
        }

        private static void AssertLegacyMatch(
            LegacyNumberBubbleMatch expected,
            bool hasSelectedCandidate,
            bool hasStrongSelectedEvidence,
            bool hasDirectSelectedIdEvidence,
            double selectedDistance,
            bool hasOtherCandidate,
            double closestOtherDistance,
            bool otherMetadataUncertain,
            double tolerance,
            string scenario)
        {
            var actual = HybridNumberingOwnershipPolicy.ClassifyLegacyBubble(
                hasSelectedCandidate,
                hasStrongSelectedEvidence,
                hasDirectSelectedIdEvidence,
                selectedDistance,
                hasOtherCandidate,
                closestOtherDistance,
                otherMetadataUncertain,
                tolerance);

            if (actual != expected)
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOwnershipPolicy) + ": " + scenario +
                    " expected " + expected + ", got " + actual + ".");
        }

        private static void TestLegacyNumberBubbleOrchestrationPolicy()
        {
            if (!HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(false, 1, false))
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOrchestrationPolicy) + ": unique legacy candidate was not adoptable.");
            if (!HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(true, 1, true))
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOrchestrationPolicy) +
                    ": a strong original anchor did not displace its owned replacement.");
            if (!HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(false, 2, true) ||
                !HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(true, 2, true))
            {
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOrchestrationPolicy) +
                    ": authoritative duplicate legacy candidates were not collapsed to one keeper.");
            }
            if (HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(true, 1, false) ||
                HybridNumberingOwnershipPolicy.CanAdoptResolvedLegacyCandidate(false, 0, true))
            {
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOrchestrationPolicy) +
                    ": weak evidence displaced a managed bubble or an empty candidate set was accepted.");
            }

            if (!HybridNumberingOwnershipPolicy.ShouldCreateMissingBubble(true, false, false) ||
                HybridNumberingOwnershipPolicy.ShouldCreateMissingBubble(true, false, true) ||
                HybridNumberingOwnershipPolicy.ShouldCreateMissingBubble(true, true, false) ||
                HybridNumberingOwnershipPolicy.ShouldCreateMissingBubble(false, false, false))
            {
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleOrchestrationPolicy) +
                    ": missing-bubble insertion was not gated by keeper/suppression state.");
            }
        }

        private static void TestLegacyNumberBubbleIdentityReconciliation()
        {
            var resolved = HybridNumberingOwnershipPolicy.ReconcileVertexIdsWithLegacyAnchors(
                new[] { 10, 74, 36 },
                new[] { 0, 36, 0 });
            AssertIds(new[] { 10, 36, 75 }, resolved, "legacy anchor must keep ID 36 at its location");

            var unchanged = HybridNumberingOwnershipPolicy.ReconcileVertexIdsWithLegacyAnchors(
                new[] { 10, 20 },
                new[] { 0, 0 });
            AssertIds(new[] { 10, 20 }, unchanged, "IDs without legacy anchors must remain stable");

            var ambiguous = HybridNumberingOwnershipPolicy.ReconcileVertexIdsWithLegacyAnchors(
                new[] { 10, 20 },
                new[] { 36, 36 });
            AssertIds(new[] { 10, 20 }, ambiguous, "a duplicated legacy ID claim must not be guessed");
        }

        private static void AssertIds(int[] expected, int[] actual, string scenario)
        {
            if (expected.Length != actual.Length)
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleIdentityReconciliation) +
                    ": " + scenario + " returned the wrong number of IDs.");

            for (int i = 0; i < expected.Length; i++)
            {
                if (expected[i] == actual[i]) continue;
                throw new InvalidOperationException(
                    nameof(TestLegacyNumberBubbleIdentityReconciliation) +
                    ": " + scenario + " expected " + expected[i] + " at index " + i +
                    ", got " + actual[i] + ".");
            }
        }

        private static void TestJtbUtmDxfScaleReader()
        {
            string path = Path.Combine(Path.GetTempPath(), "hybrid-jtb-utm-" + Guid.NewGuid().ToString("N") + ".dxf");
            try
            {
                File.WriteAllText(
                    path,
                    "0\r\nSECTION\r\n2\r\nTABLES\r\n0\r\nTABLE\r\n2\r\nUCS\r\n" +
                    "0\r\nUCS\r\n2\r\nutm\r\n146\r\n-0.999640840\r\n0\r\nENDTAB\r\n0\r\nENDSEC\r\n",
                    System.Text.Encoding.ASCII);

                if (!JtbUtmDxfScaleReader.TryReadRawScale(path, out var rawScale))
                    throw new InvalidOperationException(nameof(TestJtbUtmDxfScaleReader) + ": scale was not found.");
                AssertNear(-0.999640840, rawScale, nameof(TestJtbUtmDxfScaleReader));
                if (!JtbUtmDxfScaleReader.IsValidUtmRawScale(rawScale) ||
                    JtbUtmDxfScaleReader.IsValidUtmRawScale(Math.Abs(rawScale)) ||
                    JtbUtmDxfScaleReader.IsValidUtmRawScale(-1e-13))
                    throw new InvalidOperationException(nameof(TestJtbUtmDxfScaleReader) + ": UTM sign/magnitude policy failed.");

                File.WriteAllText(
                    path,
                    "0\r\nTABLE\r\n2\r\nUCS\r\n0\r\nUCS\r\n2\r\nother\r\n146\r\n-0.999\r\n0\r\nENDTAB\r\n",
                    System.Text.Encoding.ASCII);
                if (JtbUtmDxfScaleReader.TryReadRawScale(path, out _))
                    throw new InvalidOperationException(nameof(TestJtbUtmDxfScaleReader) + ": unrelated UCS was accepted.");
            }
            finally
            {
                try { if (File.Exists(path)) File.Delete(path); }
                catch { }
            }
        }

        private static void AssertNear(double expected, double actual, string name)
        {
            if (Math.Abs(expected - actual) > 1e-9)
                throw new InvalidOperationException(
                    name + ": expected " + expected.ToString("R", CultureInfo.InvariantCulture) +
                    ", got " + actual.ToString("R", CultureInfo.InvariantCulture) + ".");
        }
    }
}
