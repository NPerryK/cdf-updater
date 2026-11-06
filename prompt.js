/*
 * CDF Updater - prompt for Claude (claude.ai subscription)
 * Builds the text that goes to Claude with the PDFs - pasted into a chat, or saved once as the
 * instructions of a Claude Project: the instructions from ai-prompt.txt + a strict JSON output
 * format that the task pane can read back.
 *
 * Every prompt carries a short version code (a fingerprint of its text). Claude copies it into
 * each answer, so the pane can tell when a reply was made with older project instructions.
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

  function clean(base) {
    return String(base || DEFAULT_PROMPT).replace(/\r\n?/g, "\n").replace(/^﻿/, "").trim();
  }

  // FNV-1a 32-bit -> 6 hex characters. Same text -> same code on every computer.
  function fingerprint(text) {
    var h = 0x811c9dc5;
    for (var i = 0; i < text.length; i++) {
      h ^= text.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ("00000000" + h.toString(16)).slice(-8).slice(0, 6).toUpperCase();
  }

  function formatSection(keys) {
    return "OUTPUT FORMAT (this overrides any other output instruction)\n" +
      "- The user uploads one or more CDF form PDFs, each optionally with its ATT PDF. Pair each CDF with its ATT by the CDF number in the file names.\n" +
      "- Reply with exactly ONE fenced code block marked json, containing a JSON array with one object per CDF. No text before or after the block.\n" +
      "- Every object has these keys, all strings except flags:\n" +
      keys.map(function (k) { return "    " + k + ": " + FIELD_DOCS[k]; }).join("\n") + "\n" +
      "    flags: array of short strings - points the engineer should double-check (empty array if none)\n";
  }

  // returns { text, version }
  function build(base) {
    var keys = Object.keys(FIELD_DOCS);
    var body = clean(base) + "\n\n" + formatSection(keys);
    var version = fingerprint(body);
    var example = {};
    keys.forEach(function (k) { example[k] = "..."; });
    example.flags = ["..."];
    example.prompt_version = version;
    var text = body +
      "    prompt_version: always \"" + version + "\" - copy it exactly; it tells the add-in which version of these instructions you followed\n" +
      "- Example of the shape (values are placeholders):\n" +
      "```json\n" + JSON.stringify([example], null, 2) + "\n```\n";
    return { text: text, version: version };
  }

  function pastePrompt(base) { return build(base).text; }

  // Sent in the same chat when Claude answered in another format: it redoes the answer from the
  // PDFs that are already in the chat, so nothing has to be uploaded again.
  function fixupMessage(base) {
    return "Your previous answer is not in the format I need. Redo it for the same PDFs that are already in " +
      "this chat, following the instructions below exactly. Reply with only the json code block.\n\n" + build(base).text;
  }
  function promptVersion(base) { return build(base).version; }

  root.CdfPrompt = { build: build, pastePrompt: pastePrompt, promptVersion: promptVersion, fixupMessage: fixupMessage, DEFAULT_PROMPT: DEFAULT_PROMPT };
})(typeof self !== "undefined" ? self : this);
