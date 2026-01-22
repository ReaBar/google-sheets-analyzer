"""
Apps Script Manager
Manages Google Apps Script via Clasp (Node.js)
"""
import logging
import subprocess
import json
from pathlib import Path
from typing import Dict, List, Optional, Any

from .config import CLASP_CONFIG_PATH, PROJECT_ROOT

logger = logging.getLogger(__name__)


class AppsScriptManager:
    """Manages Google Apps Script using Clasp"""
    
    def __init__(self, clasp_config_path: Path = None):
        self.clasp_config_path = clasp_config_path or CLASP_CONFIG_PATH
        self.apps_script_dir = PROJECT_ROOT / "apps-script"
        self._ensure_clasp_installed()
    
    def _ensure_clasp_installed(self):
        """Check if clasp is installed, install if not"""
        try:
            result = subprocess.run(
                ["clasp", "--version"],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                logger.info(f"Clasp is installed: {result.stdout.strip()}")
                return
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        
        logger.warning("Clasp not found. Install with: npm install -g @google/clasp")
    
    def _run_clasp_command(self, command: List[str], cwd: Path = None) -> Dict[str, Any]:
        """Run a clasp command and return result"""
        if cwd is None:
            cwd = self.apps_script_dir if self.apps_script_dir.exists() else PROJECT_ROOT
        
        try:
            result = subprocess.run(
                ["clasp"] + command,
                cwd=cwd,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            return {
                "success": result.returncode == 0,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "returncode": result.returncode
            }
        except subprocess.TimeoutExpired:
            return {
                "success": False,
                "error": "Command timed out"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def login(self) -> bool:
        """Login to clasp"""
        logger.info("Logging in to clasp...")
        result = self._run_clasp_command(["login"])
        return result["success"]
    
    def clone(self, script_id: str) -> bool:
        """Clone an existing Apps Script project"""
        logger.info(f"Cloning Apps Script: {script_id}")
        
        # Ensure apps-script directory exists
        self.apps_script_dir.mkdir(parents=True, exist_ok=True)
        
        result = self._run_clasp_command(
            ["clone", script_id, "--rootDir", "apps-script"],
            cwd=PROJECT_ROOT
        )
        
        return result["success"]
    
    def create(
        self,
        title: str,
        parent_id: str,
        script_type: str = "sheets"
    ) -> bool:
        """Create a new Apps Script project"""
        logger.info(f"Creating Apps Script: {title}")
        
        # Ensure apps-script directory exists
        self.apps_script_dir.mkdir(parents=True, exist_ok=True)
        
        result = self._run_clasp_command(
            [
                "create",
                "--type", script_type,
                "--title", title,
                "--parentId", parent_id,
                "--rootDir", "apps-script"
            ],
            cwd=PROJECT_ROOT
        )
        
        return result["success"]
    
    def pull(self) -> bool:
        """Pull latest code from Google"""
        logger.info("Pulling latest Apps Script code...")
        result = self._run_clasp_command(["pull"], cwd=self.apps_script_dir)
        return result["success"]
    
    def push(self) -> bool:
        """Push local code to Google"""
        logger.info("Pushing Apps Script code to Google...")
        result = self._run_clasp_command(["push"], cwd=self.apps_script_dir)
        return result["success"]
    
    def list_files(self) -> List[str]:
        """List all files in the Apps Script project"""
        result = self._run_clasp_command(["list"], cwd=self.apps_script_dir)
        if result["success"]:
            return result["stdout"].strip().split("\n")
        return []
    
    def get_config(self) -> Optional[Dict[str, Any]]:
        """Get clasp configuration"""
        if self.clasp_config_path.exists():
            with open(self.clasp_config_path, "r") as f:
                return json.load(f)
        return None
    
    def set_config(self, config: Dict[str, Any]) -> bool:
        """Set clasp configuration"""
        try:
            with open(self.clasp_config_path, "w") as f:
                json.dump(config, f, indent=2)
            return True
        except Exception as e:
            logger.error(f"Error setting config: {e}")
            return False
