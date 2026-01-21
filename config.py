#!/usr/bin/env python3
"""
Configuration Management for DOCX to XML Pipeline

This module provides centralized configuration management for the entire
DOCX conversion pipeline.

Example Usage:
    from config import get_config, PipelineConfig

    # Get current configuration
    config = get_config()
    print(config.output_dir)
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Union


# ============================================================================
# CONFIGURATION DATACLASSES
# ============================================================================

@dataclass
class AIConfig:
    """Configuration for optional AI enhancement."""
    enabled: bool = False  # AI is optional for DOCX
    model: str = "claude-sonnet-4-20250514"
    temperature: float = 0.0  # No creativity - exact transcription
    max_tokens: int = 8192

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ExtractionConfig:
    """Configuration for DOCX content extraction."""
    extract_images: bool = True
    extract_tables: bool = True
    extract_styles: bool = True
    preserve_formatting: bool = True
    min_image_size: int = 50  # Minimum image dimension to extract

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationConfig:
    """Configuration for DTD validation."""
    dtd_path: Path = field(default_factory=lambda: Path("RITTDOCdtd/v1.1/RittDocBook.dtd"))
    max_iterations: int = 3
    generate_reports: bool = True
    report_format: str = "xlsx"  # xlsx or json

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["dtd_path"] = str(self.dtd_path)
        return d


@dataclass
class OutputConfig:
    """Configuration for output settings."""
    output_dir: Path = field(default_factory=lambda: Path("output"))
    create_docx_copy: bool = False  # Copy original DOCX to output
    create_rittdoc_zip: bool = True
    include_toc: bool = True
    toc_depth: int = 3
    cleanup_intermediate: bool = False

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["output_dir"] = str(self.output_dir)
        return d


@dataclass
class APIConfig:
    """Configuration for REST API."""
    host: str = "0.0.0.0"
    port: int = 8000
    max_concurrent_jobs: int = 5
    upload_dir: Path = field(default_factory=lambda: Path("uploads"))
    result_retention_hours: int = 24
    cors_origins: List[str] = field(default_factory=lambda: ["*"])

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["upload_dir"] = str(self.upload_dir)
        return d


@dataclass
class PipelineConfig:
    """
    Complete configuration for the DOCX to XML pipeline.

    This is the main configuration class that aggregates all sub-configurations.
    """
    # Sub-configurations
    ai: AIConfig = field(default_factory=AIConfig)
    extraction: ExtractionConfig = field(default_factory=ExtractionConfig)
    validation: ValidationConfig = field(default_factory=ValidationConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    api: APIConfig = field(default_factory=APIConfig)

    # Convenience properties
    @property
    def output_dir(self) -> Path:
        return self.output.output_dir

    @property
    def dtd_path(self) -> Path:
        return self.validation.dtd_path

    def to_dict(self) -> Dict[str, Any]:
        """Convert entire configuration to dictionary."""
        return {
            "ai": self.ai.to_dict(),
            "extraction": self.extraction.to_dict(),
            "validation": self.validation.to_dict(),
            "output": self.output.to_dict(),
            "api": self.api.to_dict(),
        }

    def to_json(self, indent: int = 2) -> str:
        """Convert configuration to JSON string."""
        return json.dumps(self.to_dict(), indent=indent)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PipelineConfig":
        """Create configuration from dictionary."""
        config = cls()

        if "ai" in data:
            config.ai = AIConfig(**data["ai"])
        if "extraction" in data:
            config.extraction = ExtractionConfig(**data["extraction"])
        if "validation" in data:
            val_data = data["validation"].copy()
            if "dtd_path" in val_data:
                val_data["dtd_path"] = Path(val_data["dtd_path"])
            config.validation = ValidationConfig(**val_data)
        if "output" in data:
            out_data = data["output"].copy()
            if "output_dir" in out_data:
                out_data["output_dir"] = Path(out_data["output_dir"])
            config.output = OutputConfig(**out_data)
        if "api" in data:
            api_data = data["api"].copy()
            if "upload_dir" in api_data:
                api_data["upload_dir"] = Path(api_data["upload_dir"])
            config.api = APIConfig(**api_data)

        return config

    @classmethod
    def from_json(cls, json_str: str) -> "PipelineConfig":
        """Create configuration from JSON string."""
        return cls.from_dict(json.loads(json_str))

    @classmethod
    def from_file(cls, path: Union[str, Path]) -> "PipelineConfig":
        """Load configuration from a JSON file."""
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Configuration file not found: {path}")

        with open(path, "r", encoding="utf-8") as f:
            return cls.from_json(f.read())

    def save(self, path: Union[str, Path]):
        """Save configuration to a JSON file."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.to_json())

    @classmethod
    def from_env(cls) -> "PipelineConfig":
        """Create configuration from environment variables."""
        config = cls()

        # AI settings
        if env_enabled := os.environ.get("DOCXTOXML_AI_ENABLED"):
            config.ai.enabled = env_enabled.lower() in ("true", "1", "yes")
        if env_model := os.environ.get("DOCXTOXML_MODEL"):
            config.ai.model = env_model

        # Output settings
        if env_output := os.environ.get("DOCXTOXML_OUTPUT_DIR"):
            config.output.output_dir = Path(env_output)
        if env_rittdoc := os.environ.get("DOCXTOXML_CREATE_RITTDOC"):
            config.output.create_rittdoc_zip = env_rittdoc.lower() in ("true", "1", "yes")

        # Validation settings
        if env_dtd := os.environ.get("DOCXTOXML_DTD_PATH"):
            config.validation.dtd_path = Path(env_dtd)

        # API settings
        if env_api_host := os.environ.get("DOCXTOXML_API_HOST"):
            config.api.host = env_api_host
        if env_api_port := os.environ.get("DOCXTOXML_API_PORT"):
            config.api.port = int(env_api_port)

        return config


# ============================================================================
# GLOBAL CONFIGURATION
# ============================================================================

_global_config: Optional[PipelineConfig] = None


def get_config() -> PipelineConfig:
    """Get the global configuration instance."""
    global _global_config
    if _global_config is None:
        _global_config = PipelineConfig.from_env()
    return _global_config


def set_config(config: PipelineConfig):
    """Set the global configuration instance."""
    global _global_config
    _global_config = config


def reset_config():
    """Reset the global configuration to default."""
    global _global_config
    _global_config = None


def load_config(path: Union[str, Path]) -> PipelineConfig:
    """Load configuration from a file and set as global."""
    config = PipelineConfig.from_file(path)
    set_config(config)
    return config


if __name__ == "__main__":
    # Print example configuration
    config = get_config()
    print("Current configuration:")
    print(config.to_json())
