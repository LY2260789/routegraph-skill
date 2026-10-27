# Reconstructing a supplied scientific figure

Inspect transparent PNGs against the intended slide background before reading labels. Black text over a transparent region can appear missing in a dark image viewer. Preserve the original file; a background-composited preview is only an inspection aid.

Separate structural diagrams from photographic insets. Rebuild the main diagram as native shapes, text and continuous routes. A source photograph can remain an explicitly identified photo inset when reproducing a composite figure; do not describe that inset as editable vector geometry. If all elements must be native, reconstruct the visible photo exterior separately. The instrument scene verifier currently rejects all pictures; it is not a mixed-figure verifier. For mixed figures, inspect OOXML and confirm that each image is an intended inset, with no rasterized main diagram.

Transcribe readable scientific labels, units and values from the reference. Do not replace a wavelength or measurement with a more familiar value from general knowledge. Record unreadable or questionable content as unresolved instead of inventing it. Distinguish a label leader, a flow edge and an optical beam; a shared junction does not establish an additional process connection.

Review rendered contact points and occlusion, not only object counts: monitor stands should meet supporting surfaces; transparent enclosure outlines should not hide beams; ring interiors should reveal the intended route. PowerPoint themes and COM state can affect fills. Inspect the saved PPTX and exported preview after setting fill visibility; do not infer transparency merely from the drawing command. Preserve deliberate occlusion and source gaps.

Use namespaced helpers such as `RG-Group` to avoid PowerShell aliases such as `Group` (`Group-Object`). For Windows PowerShell 5.1 scripts containing non-ASCII text, use UTF-8 with BOM or generate special characters explicitly. Keep each logical route as one native line or freeform path; true branches may use separate paths.

Validate three separate outcomes: content/topology against the reference, native-object structure in the saved PPTX, and the rendered visual result. A successful one-off script is a worked example, not proof that a general scene compiler supports that entire class of figures. Promote shared behavior into the library only after separating it from source-specific coordinates and labels.
