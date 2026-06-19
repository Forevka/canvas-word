// MIME types used across import/export and the collaboration backend. Single
// source so the long OOXML type string isn't retyped (and mistyped) per file.

/** OOXML WordprocessingML document (.docx). */
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
/** Portable Document Format (.pdf). */
export const PDF_MIME = "application/pdf";
/** PNG image — the default for rasterized/unknown editor media. */
export const PNG_MIME = "image/png";
/** Opaque binary fallback when a media item's type is unknown. */
export const OCTET_STREAM_MIME = "application/octet-stream";
