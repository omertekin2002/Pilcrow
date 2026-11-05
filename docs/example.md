# Example review artifact

Generated from a small test fixture in a temporary Git repository. This shows
text-diff readability, not validation in Word or Desktop.

```text
# Study results

Dr. Kaya measured 3.14 units.
The treatment did not improve the outcome.
The control was unchanged.

- Review the result.

[Table]

[Row]

[Cell]

Treatment

[Cell]

12 participants

[/Row]

[/Table]
```

After changing one sentence, ordinary `git diff` produced:

```diff
diff --git a/.pilcrow/research.docx.prose.md b/.pilcrow/research.docx.prose.md
index 7f199a2..c8bf275 100644
--- a/.pilcrow/research.docx.prose.md
+++ b/.pilcrow/research.docx.prose.md
@@ -1,7 +1,7 @@
 # Study results

 Dr. Kaya measured 3.14 units.
-The treatment improved the outcome.
+The treatment did not improve the outcome.
 The control was unchanged.

 - Review the result.
```
