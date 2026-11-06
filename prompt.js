/*
 * CDF Updater - prompt for Claude (claude.ai subscription)
 * Builds the text the user pastes into Claude together with the PDFs:
 * the instructions from ai-prompt.txt + a strict JSON output format that the task pane can read back.
 */
(function (root) {
  "use strict";

  var FIELD_DOCS = {
    main_file: "file name of the CDF form PDF exactly as uploaded",
    att_file: "file name of the ATT PDF for the same CDF exactly as uploaded, or \"\" if there is none",
    cdf_number: "CDF number exactly as printed, including revision suffix",
    form_date: "Date printed on the form, dd/mm/yyyy, or empty",
    priority: "Priority as printed (ROUTINE, URGENT, AOG...)",
    description: "UPPERCASE English summary following the style rules",
    oem_category: "OEM category (OEM STRUCTURE / OEM AVIONICS...) or empty",
    drawing_number: "Drawing / installation number referenced",
    installation: "Installation title as printed",
    code: "Short code printed on the form (e.g. C0123)",
    initiator: "Name of the person who raised / owns the CDF"
  };

  var DEFAULT_PROMPT =
    "You fill rows of an A330 P2F CDF tracking sheet from CDF form PDFs and their optional ATT attachment " +
    "PDFs. Never invent values; leave a field empty and add a flag when it is not in the documents. The " +
    "description is UPPERCASE English, 25-60 words, says what was found, where and which parts, and ends " +
    "with the engineering disposition status.";

  function pastePrompt(base) {
    var keys = Object.keys(FIELD_DOCS);
    var example = {};
    keys.forEach(function (k) { example[k] = "..."; });
    example.flags = ["..."];
    return (base || DEFAULT_PROMPT).trim() + "\n\n" +
      "OUTPUT FORMAT (this overrides any other output instruction)\n" +
      "- The user uploads one or more CDF form PDFs, each optionally with its ATT PDF. Pair each CDF with its ATT by the CDF number in the file names.\n" +
      "- Reply with exactly ONE fenced code block marked json, containing a JSON array with one object per CDF. No text before or after the block.\n" +
      "- Every object has these keys, all strings except flags:\n" +
      keys.map(function (k) { return "    " + k + ": " + FIELD_DOCS[k]; }).join("\n") + "\n" +
      "    flags: array of short strings - points the engineer should double-check (empty array if none)\n" +
      "- Example of the shape (values are placeholders):\n" +
      "```json\n" + JSON.stringify([example], null, 2) + "\n```\n";
  }

  root.CdfPrompt = { pastePrompt: pastePrompt, DEFAULT_PROMPT: DEFAULT_PROMPT };
})(typeof self !== "undefined" ? self : this);
