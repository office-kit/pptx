# Sources and licences

The storytelling and wording guidance in this skill adapts, translates and condenses rules
from the projects below. Each file was rewritten for office-kit; none is copied verbatim.
Measured figures (headline mix, seconds per slide) are the original authors' data.

| Used in                                                    | Source                                                                                                                                                                                    | Licence                    |
| ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------- |
| story-talk.md                                              | [minorun365/minorun-marp-skill](https://github.com/minorun365/minorun-marp-skill) `skills/slide-story` (commit c6b20bb)                                                                   | Apache-2.0                 |
| story-talk.md, writing-ja.md                               | [minorun365/my-claude-code-settings](https://github.com/minorun365/my-claude-code-settings) `claude/skills/writing-guide` (commit 3bab8cd)                                                | MIT, © 2026 Minoru Onda    |
| story-document.md                                          | [tyroneross/pyramid-principle](https://github.com/tyroneross/pyramid-principle) `pyramid-presentation`, `pyramid-principle-core` (commit d1343aa)                                         | Apache-2.0                 |
| story-document.md, writing-ja.md, review.md, deck-text.mjs | [carnot-tech/consulting-pptx-skill](https://github.com/carnot-tech/consulting-pptx-skill) `references/slide-rules.md`, `ai-smell-lexicon.md`, `content-review-prompt.md` (commit f50edac) | MIT, © 2026 Carnot AI Inc. |
| story-document.md, writing-ja.md, review.md                | [coji/natural-japanese](https://github.com/coji/natural-japanese) `writing-constitution.md`, `forbidden-patterns.md`, `translationese.md`, `doctypes/slide.md` (commit 9a78a42)           | MIT, © 2026 coji           |

Changes made for this skill: translated between English and Japanese, merged overlapping
rules, removed rules specific to Marp, HTML output, book publishing and the original
authors' brand design, split the rules into talk and document modes, and reimplemented
the text checks in JavaScript against `@office-kit/pptx`.

The MIT-licensed sources are provided under this notice:

> Permission is hereby granted, free of charge, to any person obtaining a copy of this
> software and associated documentation files (the "Software"), to deal in the Software
> without restriction, including without limitation the rights to use, copy, modify, merge,
> publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
> to whom the Software is furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in all copies or
> substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
> INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
> PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
> FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
> OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
> DEALINGS IN THE SOFTWARE.

The Apache-2.0 sources are used under the
[Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0); the changes are
listed above.
