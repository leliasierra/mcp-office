import re
import sys
import shutil
import tempfile
import urllib.parse
from pathlib import Path

from mcp.types import Tool, TextContent

from mcp_office.handlers.base import DocumentHandler
from mcp_office.spec import SpecManager


PPT_MASTER_SCRIPTS = (
    Path(__file__).resolve().parent.parent.parent.parent
    / "sources"
    / "ppt-master-main"
    / "skills"
    / "ppt-master"
    / "scripts"
)

PPT_MASTER_SOURCE_TO_MD = PPT_MASTER_SCRIPTS / "source_to_md"

# All 8 canvas formats from ppt-master config.py
CANVAS_FORMATS = {
    "ppt169": "PPT 16:9 (1280x720) — modern projectors",
    "ppt43": "PPT 4:3 (1024x768) — traditional projectors",
    "wechat": "WeChat header (900x383)",
    "xiaohongshu": "Xiaohongshu (1242x1660)",
    "moments": "Instagram/Square (1080x1080)",
    "story": "Vertical Story (1080x1920)",
    "banner": "Horizontal banner (1920x1080)",
    "a4": "A4 print (1240x1754)",
}

TRANSITIONS = ["fade", "push", "wipe", "split", "strips", "cover", "random", "none"]
ANIMATIONS = ["auto", "fade", "fly", "zoom", "appear", "mixed", "random", "none"]
ANIMATION_TRIGGERS = ["after-previous", "on-click", "with-previous"]

SUPPORTED_SOURCE_EXTENSIONS = {
    ".pdf": "PDF document",
    ".docx": "Word document",
    ".doc": "Word 97-2003 (requires pandoc)",
    ".odt": "OpenDocument Text (requires pandoc)",
    ".rtf": "Rich Text Format (requires pandoc)",
    ".html": "HTML page",
    ".htm": "HTML page",
    ".epub": "EPUB ebook",
    ".ipynb": "Jupyter Notebook",
    ".md": "Markdown",
    ".markdown": "Markdown",
    ".pptx": "PowerPoint presentation",
    ".pptm": "Macro-enabled PowerPoint",
    ".ppsx": "PowerPoint slideshow",
    ".ppsm": "Macro-enabled slideshow",
    ".potx": "PowerPoint template",
    ".potm": "Macro-enabled template",
    ".xlsx": "Excel workbook",
    ".xlsm": "Macro-enabled workbook",
    ".tex": "LaTeX (requires pandoc)",
    ".latex": "LaTeX (requires pandoc)",
    ".rst": "reStructuredText (requires pandoc)",
    ".org": "Emacs Org-mode (requires pandoc)",
    ".typ": "Typst (requires pandoc)",
}


class PptHandler(DocumentHandler):
    """Handler for PowerPoint operations using ppt-master engine"""

    def __init__(self):
        self._scripts_loaded = False

    def _ensure_ppt_master(self):
        if self._scripts_loaded:
            return
        scripts_dir = str(PPT_MASTER_SCRIPTS.resolve())
        if not Path(scripts_dir).exists():
            raise FileNotFoundError(
                f"ppt-master scripts not found at {scripts_dir}. "
                f"Clone ppt-master to: {PPT_MASTER_SCRIPTS.parent.parent}"
            )
        if scripts_dir not in sys.path:
            sys.path.insert(0, scripts_dir)
        if str(PPT_MASTER_SOURCE_TO_MD.resolve()) not in sys.path:
            sys.path.insert(0, str(PPT_MASTER_SOURCE_TO_MD.resolve()))
        self._scripts_loaded = True

    def get_tools(self) -> list[Tool]:
        return [
            self._ppt_generate_tool(),
            self._ppt_convert_source_tool(),
        ]

    def _ppt_generate_tool(self) -> Tool:
        fmt_desc = "Canvas format: " + ", ".join(
            f"{k} ({v})" for k, v in CANVAS_FORMATS.items()
        ) + ". Default: ppt169"
        trans_desc = "Slide transition: " + ", ".join(TRANSITIONS) + ". Default: fade"
        anim_desc = "Entrance animation: " + ", ".join(ANIMATIONS) + ". Default: auto"
        trigger_desc = "Animation trigger: " + ", ".join(ANIMATION_TRIGGERS) + ". Default: after-previous"

        return Tool(
            name="ppt_generate",
            description=(
                "Generate a native editable PPTX from SVG files using the "
                "ppt-master conversion engine. Accepts a directory of SVG "
                "files or a ppt-master project directory (with svg_output/). "
                "Returns path to the generated .pptx."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "svg_source": {
                        "type": "string",
                        "description": (
                            "Directory containing SVG slide files, or a "
                            "ppt-master project directory (with svg_output/)"
                        ),
                    },
                    "output": {
                        "type": "string",
                        "description": "Output path for the .pptx file",
                    },
                    "format": {
                        "type": "string",
                        "description": fmt_desc,
                    },
                    "transition": {
                        "type": "string",
                        "description": trans_desc,
                    },
                    "transition_duration": {
                        "type": "number",
                        "description": "Transition duration in seconds. Default: 0.5",
                    },
                    "animation": {
                        "type": "string",
                        "description": anim_desc,
                    },
                    "animation_duration": {
                        "type": "number",
                        "description": "Per-element animation duration in seconds. Default: 0.4",
                    },
                    "animation_stagger": {
                        "type": "number",
                        "description": "Delay between animated elements in seconds. Default: 0.5",
                    },
                    "animation_trigger": {
                        "type": "string",
                        "description": trigger_desc,
                    },
                    "auto_advance": {
                        "type": "number",
                        "description": "Auto-advance interval in seconds (kiosk mode). Omit for manual advance.",
                    },
                    "merge_paragraphs": {
                        "type": "boolean",
                        "description": "Merge dy-stacked paragraphs into one editable text frame. Default: true",
                    },
                    "title": {
                        "type": "string",
                        "description": "Document title metadata",
                    },
                    "author": {
                        "type": "string",
                        "description": "Document author metadata",
                    },
                },
                "required": ["svg_source", "output"],
            },
        )

    def _ppt_convert_source_tool(self) -> Tool:
        exts = ", ".join(sorted(SUPPORTED_SOURCE_EXTENSIONS.keys()))
        return Tool(
            name="ppt_convert_source",
            description=(
                "Convert a source document (PDF, DOCX, PPTX, XLSX, URL, etc.) "
                "to Markdown text using ppt-master converters. Returns the "
                "extracted Markdown content and saves it to a file."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "source": {
                        "type": "string",
                        "description": (
                            f"Path to a source file ({exts}) or a URL"
                        ),
                    },
                    "output": {
                        "type": "string",
                        "description": "Optional output path for the .md file. Auto-derived if omitted.",
                    },
                },
                "required": ["source"],
            },
        )

    async def execute(self, tool_name: str, arguments: dict) -> list[TextContent]:
        try:
            if tool_name == "ppt_generate":
                return self._generate(arguments)
            if tool_name == "ppt_convert_source":
                return self._convert_source(arguments)
            return self.error_result(f"Unknown tool: {tool_name}")
        except Exception as e:
            return self.error_result(str(e))

    # ── ppt_generate ──────────────────────────────────────────

    def _collect_svg_files(self, source: Path) -> tuple[list[Path], Path]:
        source = source.resolve()
        if not source.exists():
            raise FileNotFoundError(f"svg_source not found: {source}")
        if source.is_file():
            if source.suffix.lower() != ".svg":
                raise ValueError(f"File is not an SVG: {source}")
            return [source], source.parent
        for sub in ["svg_final", "svg_output"]:
            d = source / sub
            if d.exists():
                files = sorted(d.glob("*.svg"))
                if files:
                    return files, d
        files = sorted(source.glob("*.svg"))
        if files:
            return files, source
        raise FileNotFoundError(f"No SVG files found in: {source}")

    def _ensure_project_structure(self, svg_source: Path) -> tuple[Path, list[Path], bool]:
        svg_source = svg_source.resolve()
        # Already a project dir with svg_output/
        if (svg_source / "svg_output").exists():
            svg_files, _ = self._collect_svg_files(svg_source)
            return svg_source, svg_files, False
        # Single SVG file
        if svg_source.is_file():
            if svg_source.suffix.lower() != ".svg":
                raise ValueError(f"Not an SVG file: {svg_source}")
            tmpdir = Path(tempfile.mkdtemp(prefix="pptgen_"))
            svg_output = tmpdir / "svg_output"
            svg_output.mkdir(parents=True)
            shutil.copy2(svg_source, svg_output / svg_source.name)
            return tmpdir, [svg_output / svg_source.name], True
        # Directory: find SVG files, copy into tmp project
        svg_files, _ = self._collect_svg_files(svg_source)
        if svg_files:
            tmpdir = Path(tempfile.mkdtemp(prefix="pptgen_"))
            svg_output = tmpdir / "svg_output"
            svg_output.mkdir(parents=True)
            for f in svg_files:
                shutil.copy2(f, svg_output / f.name)
            return tmpdir, svg_files, True
        raise FileNotFoundError(f"No SVG files found in: {svg_source}")

    def _generate(self, args: dict) -> list[TextContent]:
        self._ensure_ppt_master()

        svg_source = Path(args["svg_source"])
        output_path = Path(args["output"])
        canvas_format = args.get("format", "ppt169")
        transition = args.get("transition", "fade")
        transition_duration = args.get("transition_duration", 0.5)
        animation = args.get("animation", "auto")
        animation_duration = args.get("animation_duration", 0.4)
        animation_stagger = args.get("animation_stagger", 0.5)
        animation_trigger = args.get("animation_trigger", "after-previous")
        auto_advance = args.get("auto_advance")
        merge_paragraphs = args.get("merge_paragraphs", True)
        title = args.get("title")
        author = args.get("author")

        if canvas_format not in CANVAS_FORMATS:
            valid = ", ".join(CANVAS_FORMATS.keys())
            raise ValueError(f"Unknown format '{canvas_format}'. Valid: {valid}")
        if transition not in TRANSITIONS:
            raise ValueError(f"Unknown transition '{transition}'. Valid: {', '.join(TRANSITIONS)}")
        if animation not in ANIMATIONS:
            raise ValueError(f"Unknown animation '{animation}'. Valid: {', '.join(ANIMATIONS)}")
        if animation_trigger not in ANIMATION_TRIGGERS:
            raise ValueError(f"Unknown trigger '{animation_trigger}'. Valid: {', '.join(ANIMATION_TRIGGERS)}")

        project_dir, svg_files, was_temp = self._ensure_project_structure(svg_source)

        try:
            from finalize_svg import finalize_project

            options = {
                "embed_icons": True,
                "align_images": True,
                "flatten_text": True,
                "fix_rounded": True,
            }
            finalize_project(project_dir, options, quiet=True)

            from svg_to_pptx.pptx_discovery import find_svg_files, find_notes_files
            from svg_to_pptx.pptx_builder import create_pptx_with_native_svg

            final_svgs, _ = find_svg_files(project_dir, source="final")
            if not final_svgs:
                final_svgs = svg_files

            notes = find_notes_files(project_dir, final_svgs)

            output_path.parent.mkdir(parents=True, exist_ok=True)

            doc_metadata = {}
            if title:
                doc_metadata["title"] = title
            if author:
                doc_metadata["author"] = author

            success = create_pptx_with_native_svg(
                svg_files=final_svgs,
                output_path=output_path,
                canvas_format=canvas_format,
                verbose=False,
                transition=transition,
                transition_duration=transition_duration,
                animation=animation,
                animation_duration=animation_duration,
                animation_stagger=animation_stagger,
                animation_trigger=animation_trigger,
                auto_advance=auto_advance,
                merge_paragraphs=merge_paragraphs,
                use_native_shapes=True,
                enable_notes=bool(notes),
                notes=notes,
                doc_metadata=doc_metadata or None,
            )

            if not success:
                return self.error_result("PPTX generation failed")

            SpecManager("powerpoint", str(output_path)).append({
                "type": "generate",
                "svg_source": str(svg_source),
                "format": canvas_format,
                "slides": len(final_svgs),
            })

            return self.success_result(
                f"Presentation generated: {output_path} ({len(final_svgs)} slides)"
            )
        finally:
            if was_temp:
                shutil.rmtree(project_dir, ignore_errors=True)

    # ── ppt_convert_source ────────────────────────────────────

    def _is_url(self, source: str) -> bool:
        parsed = urllib.parse.urlparse(source)
        return parsed.scheme in ("http", "https")

    def _convert_source(self, args: dict) -> list[TextContent]:
        self._ensure_ppt_master()

        source = args["source"]
        output_path = args.get("output")

        if self._is_url(source):
            return self._convert_web_url(source, output_path)

        source_path = Path(source)
        if not source_path.exists():
            raise FileNotFoundError(f"Source not found: {source_path}")

        suffix = source_path.suffix.lower()

        if suffix == ".pdf":
            return self._convert_pdf(source_path, output_path)
        if suffix in (".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm"):
            return self._convert_ppt(source_path, output_path)
        if suffix in (".xlsx", ".xlsm"):
            return self._convert_excel(source_path, output_path)
        if suffix in (".md", ".markdown"):
            content = source_path.read_text(encoding="utf-8")
            if output_path:
                Path(output_path).write_text(content, encoding="utf-8")
            return self.success_result(
                f"Read Markdown from {source_path} ({len(content)} chars)"
                + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
            )
        # doc_to_md handles: docx, html, epub, ipynb + pandoc fallbacks
        return self._convert_doc(source_path, output_path)

    def _convert_pdf(self, source: Path, output: str | None) -> list[TextContent]:
        from pdf_to_md import extract_pdf_to_markdown

        out_path = str(output) if output else None
        content = extract_pdf_to_markdown(str(source), output_path=out_path)
        if not content:
            return self.error_result(f"Failed to convert PDF: {source}")
        return self.success_result(
            f"Converted {source.name} to Markdown ({len(content)} chars)"
            + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
        )

    def _convert_ppt(self, source: Path, output: str | None) -> list[TextContent]:
        from ppt_to_md import convert_presentation_to_markdown

        out_path = str(output) if output else None
        content = convert_presentation_to_markdown(str(source), output_path=out_path)
        if not content:
            return self.error_result(f"Failed to convert PPT: {source}")
        return self.success_result(
            f"Converted {source.name} to Markdown ({len(content)} chars)"
            + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
        )

    def _convert_excel(self, source: Path, output: str | None) -> list[TextContent]:
        from excel_to_md import convert_to_markdown

        out_path = str(output) if output else None
        content = convert_to_markdown(str(source), output_path=out_path)
        if not content:
            return self.error_result(f"Failed to convert Excel: {source}")
        return self.success_result(
            f"Converted {source.name} to Markdown ({len(content)} chars)"
            + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
        )

    def _convert_doc(self, source: Path, output: str | None) -> list[TextContent]:
        from doc_to_md import convert_to_markdown

        out_path = str(output) if output else None
        content = convert_to_markdown(str(source), output_path=out_path)
        if not content:
            return self.error_result(f"Failed to convert document: {source}")
        return self.success_result(
            f"Converted {source.name} to Markdown ({len(content)} chars)"
            + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
        )

    def _convert_web_url(self, url: str, output: str | None) -> list[TextContent]:
        from web_to_md import process_url

        if output:
            ok, result_url, err = process_url(url, output_file=output)
        else:
            tmp = Path(tempfile.mktemp(suffix=".md"))
            ok, result_url, err = process_url(url, output_file=str(tmp))
            if ok:
                content = tmp.read_text(encoding="utf-8")
                tmp.unlink()
                return self.success_result(
                    f"Fetched {url} as Markdown ({len(content)} chars)"
                    + (f"\n\n{content[:5000]}" if len(content) > 5000 else f"\n\n{content}")
                )

        if not ok:
            return self.error_result(f"Failed to fetch URL: {err}")
        return self.success_result(f"Fetched and saved from URL: {url}")
