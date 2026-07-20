# Academic equation style

## Placement

- Keep an inline equation inside the grammatical sentence and preserve surrounding punctuation.
- Put a display equation in its own paragraph.
- Introduce every display equation in the preceding prose. Do not let a display appear without a textual lead-in.
- Refer to a numbered equation through its MathType reference field, for example “如式 (1) 所示”. Never type the number manually.
- Begin the first explanatory paragraph after a display with `其中，` when defining notation in Traditional Chinese.
- Begin the first explanatory sentence after an English display with `where` and keep it grammatically connected to the equation.

## Automatic classification and numbering

Scan the complete document before creating the manifest. Classification requires two passes because a later sentence can make an earlier display equation require a number.

1. Inventory every mathematical expression and every prose phrase that refers to a specific equation.
2. Classify each expression using the table below.
3. Promote any referenced display to `display_numbered`.
4. Assign sequential equation numbers in document order.
5. Replace every prose reference with a unique dynamic-reference marker targeting the numbered equation.

| Semantic class | Use when | Manifest representation |
|---|---|---|
| `inline` | The expression is a short grammatical component of a sentence, does not need its own visual line, and is not referenced as an equation elsewhere. | `layout: "inline"`, `numbered: false` |
| `display` | The expression needs emphasis, contains a derivation, fraction, matrix, cases, long operators, or would interrupt sentence readability, but no later prose needs a stable equation identifier. | `layout: "display"`, `numbered: false` |
| `display_numbered` | The display is referenced elsewhere, reused in a derivation or comparison, is a governing/model/result equation, already has a source number, or must remain addressable under the user's or publication's style. | `layout: "display"`, `numbered: true` |
| `reference` | Prose points to a specific numbered equation through wording such as `如式`, `由式`, `式中`, `Equation`, `Eq.`, `as shown in`, or an equivalent semantic reference. | One unique `references[].marker` targeting the numbered equation |

Apply these tie-breakers in order:

- Preserve an explicit user, journal, or existing-document choice unless it is structurally invalid.
- Use `inline` only when the expression remains grammatical and readable inside the sentence.
- Use `display` for a visually complex or emphasized equation even when it is not numbered.
- Number only displays that need a stable identifier. Do not number a disposable intermediate expression solely because it is displayed.
- If later prose specifically refers to an unnumbered display, promote that display to `display_numbered`; never type a static number as a shortcut.
- In PowerPoint, classify equations as floating displays only. Word-style MathType numbering and dynamic references are unavailable.

## Required paragraph pattern

Place every `display` and `display_numbered` equation in its own centered paragraph. Put a textual lead-in immediately before it. Start the following paragraph with `其中，` in Traditional Chinese, then define every symbol, index, subscript, superscript, unit, and physical meaning. Do not leave a variable used in the equation or body undefined, and do not add a variable that is absent from the equation or surrounding body. In particular, 不得列出未於正文說明的變數。

Source pattern:

```text
影像量測所得的離散角度增量可累加為總偏折角，其定義如下：

{{MATH:cumulative_deflection}}

其中，θ_y 為 y 方向之累積偏折角，單位為 rad；Δθ_(y,i) 為第 i 個取樣位置之偏折角增量，單位為 rad；i 為取樣位置索引；N 為取樣位置總數。

由式 {{EQREF:cumulative_deflection_result}} 得知，累積偏折角為各取樣位置偏折角增量之總和。
```

The reference entry for `{{EQREF:cumulative_deflection_result}}` targets `cumulative_deflection`. Use a different marker for each additional occurrence, even when several occurrences target the same equation.

Required Traditional Chinese dynamic-reference forms include:

```text
……，其於 y 方向之累積偏折角可表示如式 {{EQREF:cumulative_deflection_statement}} 所示。
……，可能缺乏直觀的可比性。因此，如式 {{EQREF:normalized_deflection_comparison}} 所示，採用了……。
由式 {{EQREF:governing_equation_derivation}} 得知……。
```

After rendering, the markers appear as dynamic references such as `如式 (2) 所示` and `由式 (1) 得知`; the visible numbers must never be typed into the source text.

## Mathematical typography contract

Apply one notation standard across the entire document. The same symbol must retain the same mathematical role and style in display equations, inline math, prose, captions, tables, and definition paragraphs.

| Role | Required style | TeX or MathML pattern |
|---|---|---|
| Scalar variable, including a Greek letter used as a variable | Italic | `x`, `\alpha`, or `<mi mathvariant="italic">x</mi>` |
| Vector | Bold lowercase | `\mathbf{x}` or `<mi mathvariant="bold">x</mi>` |
| Matrix or tensor | Bold uppercase | `\mathbf{A}` or `<mi mathvariant="bold">A</mi>` |
| Standard function or operator name | Upright Roman | `\sin`, `\exp`, `\det`, or `<mi mathvariant="normal">sin</mi>` |
| Acronym, word, or descriptive label | Upright Roman | `\mathrm{MSE}`, `\mathrm{ref}`, or `<mtext>MSE</mtext>` |
| Mathematical constant and differential symbol | Upright Roman | `\mathrm{e}`, `\mathrm{i}`, `\mathrm{d}x` |
| SI unit | Upright Roman with a space after the value | `10\ \mathrm{m\,s^{-1}}` |
| Numeric index | Upright | `<mn>1</mn>` |
| Symbolic index | Italic | `x_i` where `i` is a variable index |

Use MathType style attributes or MathML `mathvariant`; never use Unicode presentation characters to imitate mathematical bold or italic. Preserve the document's selected math font family unless the user or target publication specifies another one. Style conveys semantics and must not depend on a particular installed body-text font.

For every document:

1. Inventory each symbol and assign one role before creating the manifest.
2. Encode that role consistently in every TeX or MathML expression.
3. Match the same style when the symbol appears in surrounding prose or definitions.
4. Keep multi-letter functions, abbreviations, textual conditions, and descriptive subscripts upright.
5. Keep scalar variables and variable indices italic, including Greek letters used as variables.
6. Keep vectors bold lowercase and matrices or tensors bold uppercase unless the user's established notation explicitly differs.
7. Visually inspect the complete output because structural OLE validation cannot prove font weight, slant, or consistency.

These rules follow the IEEE guidance that variables remain italic in both prose and equations, vectors use bold type, and function names use Roman type. Consult the [IEEE Mathematics Style Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/Editing-Mathematics.pdf) and [IEEE Math Typesetting Guide](https://journals.ieeeauthorcenter.ieee.org/wp-content/uploads/sites/7/IEEE-Math-Typesetting-Guide-for-LaTeX-Users.pdf) when the target publication does not provide a stricter house style.

## Academic English dynamic-reference templates

The `{{EQREF:...}}` token below is a unique source-occurrence marker. Replace it through MathType so the rendered `(n)` remains dynamic; never leave or manually type the number.

```text
..., and its cumulative deflection angle in the y direction can be expressed as shown in Eq. {{EQREF:cumulative_deflection}}.
..., which may lack an intuitive basis for comparison. Therefore, as shown in Eq. {{EQREF:normalized_deflection}}, ... was adopted.
Equation {{EQREF:governing_equation}} indicates that ...
```

Use `Eq.` for a singular reference and `Eqs.` for plural references unless the target journal specifies another house style.

## Definitions

Define all non-obvious notation at first use:

- Roman and Greek symbols.
- Vectors, matrices, tensors, and sets.
- Subscripts and superscripts.
- Summation/product indices and their bounds.
- Operators whose meaning is domain-specific.
- Units and dimensions.
- Physical, statistical, or engineering meaning.

Apply a bidirectional completeness check: every symbol in the equation must be defined or already unambiguous from the immediately preceding prose, and the definition paragraph must not introduce symbols absent from the equation. Include every index, bound, superscript, subscript, and unit exactly once at first use.

Example:

```text
系統的離散時間狀態方程式可表示為：

{{MATH:state_transition}}

其中，x_k 為第 k 個取樣時刻的狀態向量，A 為狀態轉移矩陣，B 為輸入矩陣，u_k 為控制輸入，w_k 為程序雜訊；時間索引 k 為無因次整數。
```

English example:

```text
The discrete-time state equation is expressed as follows:

{{MATH:state_transition}}

where x_k is the state vector at sampling instant k, A is the state-transition matrix, B is the input matrix, u_k is the control-input vector, and w_k is the process-noise vector; k is a dimensionless integer time index.
```

## TeX guidance for MathType Toggle TeX

Prefer portable constructs:

- `\frac{a}{b}`
- `x_i`, `x^{(k)}`
- `\sum_{i=1}^{n}`
- `\sqrt{x}`
- `\mathbf{x}`, `\boldsymbol{\theta}` when supported by the installed translator
- `\left( ... \right)`
- `\begin{matrix} ... \end{matrix}` only after a local conversion test

Avoid dependencies on:

- `\documentclass`, preambles, or `\usepackage`.
- Custom macros and commands.
- `\label`, `\ref`, `\eqref`, or automatic LaTeX numbering.
- Environment features not supported by MathType's current TeX translator.
- Multiple equations encoded into one string when separate semantic objects are needed.

JSON doubles each backslash. TeX `\sigma^2` becomes JSON `"\\sigma^2"`.

## Quality gate

Structural validation proves object and field type, not appearance. Visually inspect:

- Fractions and radicals are not clipped.
- Matrices have the intended row/column count and delimiters.
- Accents attach to the intended symbol.
- Scripts have correct scope.
- Operators and differential symbols use consistent style.
- Scalar variables are italic; vectors are bold lowercase; matrices and tensors are bold uppercase.
- Functions, acronyms, descriptive labels, mathematical constants, differentials, and SI units are upright.
- Each symbol has the same role and style in the equation, prose, captions, and definition paragraph.
- A display equation and its number share the intended line and tab alignment.
- Long equations do not collide with `(n)` or overflow the text width.
