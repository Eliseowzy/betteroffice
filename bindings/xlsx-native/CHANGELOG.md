# @betteroffice/xlsx-native

## 0.6.0

### Minor Changes

- 702b174: The calculation engine now evaluates RATE, IRR, MIRR, XNPV, XIRR, IPMT, PPMT, CUMIPMT, CUMPRINC, ISPMT, DB, DDB, VDB, SLN, SYD, EFFECT, NOMINAL, PDURATION, RRI, FVSCHEDULE, DOLLARDE, DOLLARFR and CELL. RATE follows Excel's own iteration, so results match Excel's saved values.
- dac1245: The calculation engine now evaluates GAMMA, GAMMALN and the gamma, beta, normal, lognormal, chi-square, t and F distribution functions with their inverses and compatibility names, ERF and ERFC, PHI, GAUSS, CHISQ.TEST, T.TEST, F.TEST and the CONFIDENCE functions.
- 317d85c: The calculation engine now evaluates Excel's database functions DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR and DVARP, including OR rows, repeated columns, wildcards and computed criteria.
- f92d560: The calculation engine now evaluates the binomial, negative binomial, hypergeometric, Poisson, exponential and Weibull distribution functions, FISHER, STANDARDIZE, PERMUT, PEARSON, STEYX, SKEW.P, the A-variance functions, the exclusive percentile and rank functions, MODE.MULT, PROB, Z.TEST, GROWTH and LOGEST.
- c06b68d: The calculation engine now evaluates Excel's engineering functions: base conversions (BIN2DEC through OCT2HEX), the BIT functions, DELTA, GESTEP, COMPLEX, every IM complex-number function and the four Bessel functions. Complex results are written as Excel writes them.
- 74407d2: The calculation engine now evaluates more information and text functions: ERROR.TYPE, INFO, ISFORMULA, ISNONTEXT, ISOMITTED, ISREF, SHEET, SHEETS, TYPE, UNICHAR, UNICODE, DOLLAR, VALUETOTEXT, ARRAYTOTEXT, ASC, JIS, DBCS, BAHTTEXT, PHONETIC, ENCODEURL and the byte-count text functions such as LENB.
- e5bb91a: The calculation engine now evaluates 33 more math and trig functions, including the hyperbolic and reciprocal trig functions, ROMAN, ARABIC, BASE, DECIMAL, CEILING.MATH, FLOOR.MATH, MDETERM, MINVERSE, MUNIT, RANDARRAY, the SUMX2 family, PERCENTOF and DAYS360.
- 2dbf721: The calculation engine now evaluates Excel's securities functions: the COUP functions, ACCRINT, ACCRINTM, PRICE, YIELD and their discount and at-maturity variants, DISC, INTRATE, RECEIVED, DURATION, MDURATION, the TBILL functions, the odd-period functions and AMORDEGRC and AMORLINC.
- 3adc24b: `renderPng` draws in the native engine through the new raw `renderPngTarget(targetJson, optionsJson)`, which honours scale and maxWidth and maxHeight for viewports and ranges. The raw `renderPng(viewportJson)` and `renderRangePng(args)` are removed.
- 2c28650: After a recalculation, saved files carry the recalculated values: the new `WorkbookHandle.recalculate()`, or opening a file that asks for a full calculation on load. Formulas the engine cannot evaluate keep the result the file stored. Without either, saved values stay byte for byte.
- c1525f7: Formulas using functions that cannot be computed offline (CUBE functions, WEBSERVICE, RTD, STOCKHISTORY, IMAGE, add-in calls and similar), and formulas reading them, keep the value Excel saved, even inside IFERROR, and the calculation report lists them as `unevaluatedCells` in JS and `unevaluated` in Python.

### Patch Changes

- ad45f86: Adds opt-in `decodeFontFace` for packed fonts. Format package roots return a promise of `font` and optional `payload` ArrayBuffers; native bindings return byte arrays synchronously. Invalid options and decode errors reject or throw without retrying another decoder.
- b0d5d34: The native xlsx engine has the same 4 GiB memory budget per thread as the WebAssembly engine. A call past the budget fails the worker session with the typed out-of-memory error, and a failed workbook handle refuses later calls except dispose().
- 1af9969: Adds prebuilt binaries for musl Linux on x64 and arm64, such as Alpine, and for Windows on arm64.
- ec04baa: The native packages require Node.js 20.16+ or 22.3+, which they need to read installed font packs from disk.
- 681b02e: A Rust panic in a native engine now fails its session as a trap, as on WebAssembly, instead of aborting the process. The engine object that panicked refuses every later call except free.
- 6378188: Worker sessions now run on Node.js >= 20.16 / 22.3 through `worker_threads`, keeping the process alive until disposed. DOCX worker editing needs `OffscreenCanvas` and fails with a clear error without it (Node.js, Bun, Deno); open, read and save still work.
- 7cf70f8: Worker entries no longer import the module that constructs their own worker, so Next.js (Turbopack) production builds of apps that use `./worker` finish instead of hanging.
- 55c25ba: Worker files and main-thread entries import no private `#` specifiers, so worker sessions open from a CDN or with an import map, without a bundler. On Node.js, a worker thread that fails to start now fails the session instead of throwing synchronously.
- f77d138: Date functions in workbooks that use the 1904 date system now return 1904 serials: `DATE(2026,11,1)` is 44865, as in Excel. `nowSerial` is always a 1900-system serial, so TODAY(), NOW() and typed dates without a year match the calendar in 1904 workbooks too.
- f25a90c: Calculation matches Excel at numeric limits: numbers become text with 15 significant digits (`""&(0.1+0.2)` is "0.3"), formula literals keep 15 digits within Excel's range, subnormal results are 0, and MOD, MULTINOMIAL, BESSELK, the TBILL, T and F functions follow Excel's limits.
- 4990e8d: Native XLSX rendering draws charts much faster: a chart sheet that took about 400 ms now takes about 20 ms. Pixel output is unchanged.
- 4b01734: Editing many defined names in one `applyOps` call now takes time and memory linear in the number of names. Re-committing all 66,000 names of a large workbook used to exhaust memory.
- 14b82ff: Shown and returned formula text drops Excel's storage prefixes (`_xlfn.`, `_xlws.`, `_xlpm.`, `_xleta.`): the formula bar, `readCells`, structured export, defined names, conditional formats, agent reads and Python `formula()` give `=SEQUENCE(2,2)`. Saved files keep the prefixes; re-committing a shown formula leaves the cell unchanged.
- 9d433be: Formulas written in BetterOffice that use newer Excel functions (XLOOKUP, IFS, LET, LAMBDA, FILTER and others) are saved with the `_xlfn.` prefixes Excel expects, so Excel no longer opens them as #NAME?. Unedited formulas keep their exact text.
- 6d749ac: PNG export uses the installed Latin, international and CJK font packs by default, reading faces from disk with nothing fetched.
- 7a91087: Now exposes exactly the `@betteroffice/xlsx` API (same entry points, exports and types) on the native engine for Node.js, Bun and Deno; pick the engine by import. The `browser` condition throws. The `openWorkbook(Buffer)`/`Workbook` facade is removed.
- e200f5e: Speeds up style patches and number format changes on bounded ranges in standalone workbooks. Enabled by default, with the same saved files, updates and undo behavior.
- Updated dependencies [de8f092]
- Updated dependencies [d60b8e4]
- Updated dependencies [d60b8e4]
- Updated dependencies [df33f6a]
- Updated dependencies [37199d9]
- Updated dependencies [e6d5c52]
  - @betteroffice/fonts-cjk@1.0.0
  - @betteroffice/fonts-intl@1.0.0
  - @betteroffice/fonts-latin@1.0.0

## 0.5.1

## 0.5.0

### Patch Changes

- c688743: Adds prebuilt Rust engines for Node.js, Bun and Deno that open, edit, render PNGs and save documents. Image decoding stays within the render budget.
