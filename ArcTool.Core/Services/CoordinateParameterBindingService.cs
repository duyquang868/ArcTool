#nullable enable
using System;
using System.Collections.Generic;
using ArcTool.Core.Models;
using Autodesk.Revit.DB;

namespace ArcTool.Core.Services
{
    /// <summary>
    /// Registers and repairs the AT_CoordX / AT_CoordY / AT_CoordZ instance bindings for explicit coordinate categories.
    /// Keeping category selection at the caller boundary lets Element Type and Detail Type registration stay separate.
    /// </summary>
    public static class CoordinateParameterBindingService
    {
        private static readonly string[] CoordinateParameterNames =
        {
            CoordParamNames.CoordX,
            CoordParamNames.CoordY,
            CoordParamNames.CoordZ
        };

        /// <summary>
        /// Returns true when AT_CoordX is currently bound as an instance parameter to the supplied category.
        /// This is the runtime source of truth for whether a category was registered by ArcTool.
        /// </summary>
        /// <param name="doc">Active Revit document whose binding map should be inspected.</param>
        /// <param name="targetCategory">Category to test.</param>
        /// <returns>True when the coordinate parameter binding includes the category; otherwise false.</returns>
        public static bool IsCoordinateCategoryRegistered(Document doc, BuiltInCategory targetCategory)
        {
            if (doc == null)
            {
                return false;
            }

            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                if (!string.Equals(iterator.Key?.Name, CoordParamNames.CoordX, StringComparison.Ordinal))
                {
                    continue;
                }

                if (iterator.Current is not InstanceBinding instanceBinding)
                {
                    return false;
                }

                foreach (Category category in instanceBinding.Categories)
                {
                    if (category?.BuiltInCategory == targetCategory)
                    {
                        return true;
                    }
                }

                return false;
            }

            return false;
        }

        /// <summary>
        /// Ensures all coordinate value parameters exist and are bound as instance parameters to the supplied categories.
        /// Must be called inside an active Revit transaction.
        /// </summary>
        /// <param name="doc">Active Revit document that will receive the bindings.</param>
        /// <param name="group">Shared-parameter definition group that stores ArcTool coordinate definitions.</param>
        /// <param name="targetCategories">Exact categories that should receive coordinate instance parameters.</param>
        /// <param name="categoryLabel">Human-readable category label for failure messages.</param>
        public static void EnsureCoordinateParameters(
            Document doc,
            DefinitionGroup group,
            IReadOnlyList<BuiltInCategory> targetCategories,
            string categoryLabel)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (group == null)
            {
                throw new ArgumentNullException(nameof(group));
            }

            if (targetCategories == null || targetCategories.Count == 0)
            {
                throw new ArgumentException("At least one target category is required.", nameof(targetCategories));
            }

            foreach (string parameterName in CoordinateParameterNames)
            {
                EnsureParameterBinding(doc, group, parameterName, targetCategories, categoryLabel);
            }
        }

        private static void EnsureParameterBinding(
            Document doc,
            DefinitionGroup group,
            string paramName,
            IReadOnlyList<BuiltInCategory> targetCategories,
            string categoryLabel)
        {
            Definition definition = group.Definitions.get_Item(paramName);
            if (definition == null)
            {
                var options = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.Number)
                {
                    Visible = true
                };

                definition = group.Definitions.Create(options);
            }

            if (definition == null)
            {
                throw new InvalidOperationException($"Failed to create or retrieve shared parameter definition '{paramName}'.");
            }

            if (IsAlreadyBound(doc, paramName, targetCategories))
            {
                return;
            }

            if (TryReinsertMergedBinding(doc, definition, paramName, targetCategories))
            {
                return;
            }

            RegisterNewBinding(doc, definition, paramName, targetCategories, categoryLabel);
        }

        private static bool IsAlreadyBound(Document doc, string paramName, IReadOnlyList<BuiltInCategory> targetCategories)
        {
            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                if (!string.Equals(iterator.Key?.Name, paramName, StringComparison.Ordinal))
                {
                    continue;
                }

                if (iterator.Current is not InstanceBinding instanceBinding)
                {
                    return false;
                }

                return ContainsAllTargetCategories(instanceBinding.Categories, targetCategories);
            }

            return false;
        }

        private static bool TryReinsertMergedBinding(
            Document doc,
            Definition definition,
            string paramName,
            IReadOnlyList<BuiltInCategory> targetCategories)
        {
            DefinitionBindingMapIterator iterator = doc.ParameterBindings.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                if (!string.Equals(iterator.Key?.Name, paramName, StringComparison.Ordinal))
                {
                    continue;
                }

                if (iterator.Current is not InstanceBinding existingBinding)
                {
                    throw new InvalidOperationException($"Parameter '{paramName}' exists but is not an instance binding.");
                }

                CategorySet mergedCategories = doc.Application.Create.NewCategorySet();
                foreach (Category category in existingBinding.Categories)
                {
                    if (category != null)
                    {
                        mergedCategories.Insert(category);
                    }
                }

                if (ContainsAllTargetCategories(existingBinding.Categories, targetCategories))
                {
                    return true;
                }

                InsertTargetCategories(doc, mergedCategories, targetCategories);
                InstanceBinding mergedBinding = doc.Application.Create.NewInstanceBinding(mergedCategories);
                if (!doc.ParameterBindings.ReInsert(definition, mergedBinding))
                {
                    throw new InvalidOperationException($"Failed to update the existing binding for '{paramName}'.");
                }

                return true;
            }

            return false;
        }

        private static void RegisterNewBinding(
            Document doc,
            Definition definition,
            string paramName,
            IReadOnlyList<BuiltInCategory> targetCategories,
            string categoryLabel)
        {
            CategorySet categorySet = doc.Application.Create.NewCategorySet();
            InsertTargetCategories(doc, categorySet, targetCategories);

            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
            if (!doc.ParameterBindings.Insert(definition, binding))
            {
                throw new InvalidOperationException($"Failed to bind '{paramName}' to {categoryLabel}.");
            }
        }

        private static void InsertTargetCategories(
            Document doc,
            CategorySet categorySet,
            IReadOnlyList<BuiltInCategory> targetCategories)
        {
            foreach (BuiltInCategory targetCategoryId in targetCategories)
            {
                Category targetCategory = doc.Settings.Categories.get_Item(targetCategoryId);
                if (targetCategory == null)
                {
                    throw new InvalidOperationException($"Could not resolve supported coordinate category '{targetCategoryId}'.");
                }

                categorySet.Insert(targetCategory);
            }
        }

        private static bool ContainsAllTargetCategories(CategorySet categories, IReadOnlyList<BuiltInCategory> targetCategories)
        {
            int foundCategoryCount = 0;

            foreach (BuiltInCategory targetCategory in targetCategories)
            {
                bool found = false;
                foreach (Category category in categories)
                {
                    if (category?.BuiltInCategory == targetCategory)
                    {
                        found = true;
                        break;
                    }
                }

                if (found)
                {
                    foundCategoryCount++;
                }
            }

            return foundCategoryCount == targetCategories.Count;
        }
    }
}
