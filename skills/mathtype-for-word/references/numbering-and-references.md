# MathType-native numbering and references

## Contents

1. Required default
2. Native number anatomy
3. Native reference anatomy
4. Why Word numbering is not equivalent
5. Editing and field updates
6. Default scope

## Required default

The bundled `config/defaults.json` expresses the requested Format Equation Numbers state:

| Setting | Value |
|---|---|
| Mode | Simple Format |
| Chapter Number | Off |
| Section Number | Off |
| Equation Number | On, Arabic |
| Enclosure | Parentheses |
| Separator | Off |
| Change format for | New equation numbers and Whole document |
| Update automatically | On |
| Warn on first equation number | On |
| Warn on references | Off |
| Default for new documents | On in this agent workflow |

Every render normalizes the visible number to `(1)`, `(2)`, and so on, even if the interactive MathType dialog previously used a section-aware form such as `(1.1)`.

## Native number anatomy

MathType 7 uses nested Word fields. The important visible-number structure is conceptually:

```text
MACROBUTTON MTPlaceRef
  SEQ MTEqn \h
  (
  SEQ MTEqn \c \* Arabic
  )
```

The hidden sequence field advances the MathType equation counter. The current-value field displays it. The outer `MTPlaceRef` macrobutton is the target a user clicks after inserting a reference placeholder.

MathType may also insert a document setup macrobutton containing hidden reset fields. Preserve it. Do not flatten or unlink these fields.

The bridge calls `MTCommand_InsertEqnNum` first, so MathType creates its full native structure. It then removes only current chapter or section components and their separator from the visible `MTPlaceRef` field. It does not replace the native mechanism.

## Native reference anatomy

The interactive and automated workflows are the same:

1. `MTCommand_InsertEqnRef` inserts `equation reference goes here` and creates the temporary `MTReference` bookmark.
2. The target equation number's `MTPlaceRef` action runs.
3. MathType creates a target bookmark named like `ZEqnNum394416`.
4. The placeholder becomes a `GOTOBUTTON` field containing a nested `REF` field.

Conceptual field:

```text
GOTOBUTTON ZEqnNum394416
  REF ZEqnNum394416 \* Charformat \! \* MERGEFORMAT
```

This is why a literal `(1)` or a Word list cross-reference is not acceptable. It lacks MathType's target placement behavior and bookmark relationship.

## Why Word numbering is not equivalent

Do not use these for the primary workflow:

- Numbered lists.
- Captions labeled Equation.
- Manually authored `SEQ Equation` fields.
- Manually typed numbers.
- `\label` and `\ref` text imported as TeX.

They may look similar but are not MathType's numbering/reference implementation. Use them only after a probe failure and explicit user acceptance of a fallback.

## Editing and field updates

After adding, deleting, or moving numbered equations:

1. Update all Word fields.
2. Verify the `MTEqn` current values are sequential from 1.
3. Verify every `GOTOBUTTON` target bookmark exists.
4. Search for `Error! Reference source not found.`
5. Validate again with the original manifest or an updated manifest.

Never edit `ZEqnNum...` names manually. MathType owns those bookmarks.

## Default scope

`configure_mathtype_word_defaults` writes the packaged default profile to `%APPDATA%\MathTypeForWordAgent\defaults.json` and sets MathType's user warning preferences. Every document processed by this bridge uses the profile and is normalized to the requested number form.

The skill does not automate the localized Format Equation Numbers dialog with keystrokes. That would be timing- and language-dependent. Documents created manually outside this workflow retain the user's interactive MathType configuration unless the user also selects the same settings in the dialog.
