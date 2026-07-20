# PowerPoint workflow

## Supported result

Create editable `Equation.DSMT4` OLE equations directly in `.pptx`. The bridge follows MathType's PowerPoint design: PowerPoint creates and activates the OLE object, the embedded MathType editor receives bare Presentation MathML from the Windows Unicode clipboard, and MathType's `RunForConversion` verb refreshes PowerPoint's cached preview. The bridge then reads MathML back from the OLE `IDataObject` and compares it with the manifest. Word is not used as a conversion intermediary. MathType 7 inserts external objects as floating shapes because PowerPoint does not support inline external objects.

MathType's Word-only `MTPlaceRef`, `SEQ MTEqn`, `MTReference`, and `GOTOBUTTON/REF` workflow does not exist in PowerPoint. Do not emulate it with typed equation numbers or fake hyperlinks.

This is desktop UI automation. Keep the Windows session unlocked, do not use the keyboard or replace the clipboard while a formula is being committed, and process equations serially. The bridge validates the returned MathML before it saves the PPTX.

## Presentation manifest schema v1

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "energy_relation",
      "marker": "{{MATH:energy_relation}}",
      "mathml": "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>",
      "height_points": 32
    }
  ]
}
```

Rules:

- Use unique stable ASCII IDs.
- Supply one complete Presentation MathML `<math>` element. Generate it directly from the intended formula and verify every operator, script, fraction, radical, accent, and identifier before rendering.
- Put each marker alone in one top-level text box. Grouped markers and markers embedded in prose are rejected.
- Set `height_points` between 12 and 200. Omit it to use 32 points.
- Keep descriptive prose in separate text boxes above and below the marker.

## Direct bridge commands

```powershell
# Resolve these variables from the user's attached or repository files; never reuse a sample path literally.
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action probe-pptx

pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action render-pptx `
  -InputPath $inputPptx `
  -OutputPath $outputPptx `
  -ManifestPath $manifestJson

pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action validate-pptx `
  -InputPath $outputPptx `
  -ManifestPath $manifestJson
```

## Validation gate

Require all of the following before claiming completion:

- One named `Equation.DSMT4` OLE object for every manifest equation.
- Object name `MathType_<id>` and alternative text identifying the equation.
- Horizontal centering within one point.
- No unresolved `{{MATH:...}}` marker.
- Full-size rendered slide inspection with no clipping, overlap, or illegible equation sizing.
- No Word process is created by the PPTX render action.
