-- Per-document custom-font configuration. Stored at create time (the editor's
-- resolved `fonts` option: custom families with face URLs + sizing, plus toolbar
-- disables). Server-side export reads it, fetches the URLs (disk-cached), and
-- embeds the faces so a headless render matches the editor.

ALTER TABLE documents ADD COLUMN fonts_config jsonb;
