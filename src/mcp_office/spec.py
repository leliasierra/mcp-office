import json
from pathlib import Path


def spec_path_for(document_path: str) -> str:
    return str(Path(document_path).with_suffix(Path(document_path).suffix + ".spec.json"))


class SpecManager:
    @staticmethod
    def spec_path_for(document_path: str) -> str:
        return str(Path(document_path).with_suffix(Path(document_path).suffix + ".spec.json"))

    def __init__(self, document_type: str, document_path: str):
        self.document_type = document_type
        self.document_path = document_path
        self.spec_path = self.spec_path_for(document_path)
        self._spec = self._load_or_create()

    def _load_or_create(self) -> dict:
        path = Path(self.spec_path)
        if path.exists():
            try:
                with open(path) as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError):
                pass
        return {
            "type": self.document_type,
            "path": self.document_path,
            "properties": {},
            "content": [],
            "post_processing": [],
        }

    def set_property(self, key: str, value):
        self._spec["properties"][key] = value
        self._save()

    def append(self, item: dict):
        self._spec["content"].append(item)
        self._save()

    def add_post_processing(self, action: str, **kwargs):
        entry = {"action": action, **kwargs}
        self._spec["post_processing"].append(entry)
        self._save()

    @property
    def spec(self) -> dict:
        return self._spec

    def _save(self):
        Path(self.spec_path).parent.mkdir(parents=True, exist_ok=True)
        with open(self.spec_path, "w") as f:
            json.dump(self._spec, f, indent=2, ensure_ascii=False)

    def remove_last(self, item_type: str = None):
        if item_type:
            for i in range(len(self._spec["content"]) - 1, -1, -1):
                if self._spec["content"][i].get("type") == item_type:
                    self._spec["content"].pop(i)
                    self._save()
                    return
        else:
            if self._spec["content"]:
                self._spec["content"].pop()
                self._save()
