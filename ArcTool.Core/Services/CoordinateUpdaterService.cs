#nullable enable
using System;
using System.Linq;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Manages document-scoped lifecycle registration for the ArcTool Coordinate updater.
    /// </summary>
    public static class CoordinateUpdaterService
    {
        /// <summary>
        /// Stable GUID for this updater. Must never change between Revit sessions.
        /// Changing this GUID breaks updater persistence and registration state.
        /// Generated once; treat as a project artifact.
        /// </summary>
        public static readonly Guid UpdaterGuid =
            new Guid("42D71F93-3A10-41C2-86B4-0C8762BA77E9");
        // Replace with the project-approved deployment GUID before first deployment.
        // After deployment, this value must remain stable and must never be regenerated.

        /// <summary>
        /// Register the CoordinateUpdater for a specific document.
        /// Safe to call multiple times — guards against double-registration.
        /// Skips silently if AT_CoordX is not bound to the supported coordinate categories.
        /// Writes one journal comment regardless of outcome for traceability.
        /// </summary>
        public static void RegisterForDocument(Document doc, AddInId addInId)
        {
            if (doc == null)
            {
                return;
            }

            if (addInId == null)
            {
                doc.Application.WriteJournalComment(
                    "[ArcTool CoordinateUpdaterService] AddInId is null — updater registration skipped.",
                    false);
                return;
            }

            try
            {
                UpdaterId updaterId = new UpdaterId(addInId, UpdaterGuid);
                if (UpdaterRegistry.IsUpdaterRegistered(updaterId, doc))
                {
                    UpdaterRegistry.UnregisterUpdater(updaterId, doc);
                }

                var registeredFilters = CoordinateExtractionService.GetRegisteredTriggerFilters(doc);
                if (registeredFilters.Count == 0)
                {
                    doc.Application.WriteJournalComment(
                        "[ArcTool CoordinateUpdaterService] No registered coordinate categories found — updater registration skipped.",
                        false);
                    return;
                }

                CoordinateUpdater updater = new CoordinateUpdater(addInId);
                UpdaterRegistry.RegisterUpdater(updater, doc, false);

                foreach (CoordTriggerFilter triggerFilter in registeredFilters)
                {
                    UpdaterRegistry.AddTrigger(
                        updaterId,
                        doc,
                        BuildSupportedCategoryFilter(triggerFilter),
                        Element.GetChangeTypeGeometry());
                }

                string categoryList = string.Join(", ", registeredFilters.Select(CoordV1Scope.GetCategoryLabel));
                doc.Application.WriteJournalComment(
                    $"[ArcTool CoordinateUpdaterService] Updater registered successfully for: {categoryList}.",
                    false);
            }
            catch (Exception ex)
            {
                doc.Application.WriteJournalComment(
                    $"[ArcTool CoordinateUpdaterService] Registration failed: {ex.Message}",
                    false);
            }
        }

        private static FamilyInstance? FindFirstSupportedFamilyInstance(Document doc, CoordTriggerFilter triggerFilter)
        {
            BuiltInCategory category = CoordV1Scope.GetCategory(triggerFilter);
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(category)
                .Cast<FamilyInstance>()
                .FirstOrDefault(instance => triggerFilter != CoordTriggerFilter.DetailItems
                    || CoordinateDetailItemRegistryService.IsRegisteredType(doc, instance));
        }

        private static ElementFilter BuildSupportedCategoryFilter(CoordTriggerFilter triggerFilter)
        {
            return new ElementCategoryFilter(CoordV1Scope.GetCategory(triggerFilter));
        }

        /// <summary>
        /// Unregister the CoordinateUpdater for a specific document.
        /// Safe to call even if the updater was never registered.
        /// </summary>
        public static void UnregisterForDocument(Document doc, AddInId addInId)
        {
            if (doc == null)
            {
                return;
            }

            if (addInId == null)
            {
                doc.Application.WriteJournalComment(
                    "[ArcTool CoordinateUpdaterService] AddInId is null — updater unregister skipped.",
                    false);
                return;
            }

            try
            {
                UpdaterId updaterId = new UpdaterId(addInId, UpdaterGuid);
                if (UpdaterRegistry.IsUpdaterRegistered(updaterId, doc))
                {
                    UpdaterRegistry.UnregisterUpdater(updaterId, doc);
                }
            }
            catch (Exception ex)
            {
                doc.Application.WriteJournalComment(
                    $"[ArcTool CoordinateUpdaterService] Unregister failed: {ex.Message}",
                    false);
            }
        }

        /// <summary>
        /// Returns true if the updater is currently registered for the given document.
        /// </summary>
        public static bool IsRegisteredForDocument(Document doc, AddInId addInId)
        {
            if (doc == null || addInId == null)
            {
                return false;
            }

            UpdaterId updaterId = new UpdaterId(addInId, UpdaterGuid);
            return UpdaterRegistry.IsUpdaterRegistered(updaterId, doc);
        }
    }
}
