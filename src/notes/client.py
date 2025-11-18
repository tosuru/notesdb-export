
from __future__ import annotations
from typing import Iterable, Optional, List
import os
import traceback
import logging

logger = logging.getLogger(__name__)


class NotesClient:
    def __init__(self, server: str, db_file: str):
        self.server = server or ""
        self.db_file = db_file
        self.session = None
        self.db = None

    def connect(self):
        logger.debug("[client] 1. Attempting to connect to Notes...")
        try:
            import win32com.client as win32
        except Exception as e:
            logger.debug("[client] X. Failed to import win32com.client")
            raise ImportError(
                "win32com.client is required to connect to Notes/Domino.") from e

        try:
            self.session = win32.Dispatch("Lotus.NotesSession")
            # Initialize() is crucial for Lotus.NotesSession
            password = os.environ.get("NOTES_PASSWORD")
            logger.debug("[client] 2. Initializing session...")
            # パスワードは環境変数を使う。無ければinput()で聞く。
            if password:
                logger.debug(
                    "[client] Found password in NOTES_PASSWORD environment variable.")
            else:
                logger.debug(
                    "[client] NOTES_PASSWORD environment variable not found or is empty. Falling back to prompt.")
                import getpass
                password = getpass.getpass(
                    prompt="Enter your Lotus Notes password: ")
            self.session.Initialize(password)
            logger.debug(
                f"[client] 3. NotesSession initialized. User: {self.session.UserName}")

            self.db = self.session.GetDatabase(self.server, self.db_file)
            if not self.db.IsOpen:
                logger.debug(
                    f"[client] 4. Database is not open. Opening now... Server: '{self.server}', DB: '{self.db_file}'")
                self.db.Open()
            else:
                logger.debug("[client] 4. Database is already open.")
            logger.debug(
                f"[client] 5. Successfully connected to DB: {self.db.Title}")
        except Exception as e:
            logger.exception(f"[client] X. An error occurred during connect: {e}")
            raise
        return self

    def get_document_by_unid(self, unid: str):
        logger.debug(f"[client] Getting document by UNID: {unid}")
        try:
            doc = self.db.GetDocumentByUNID(unid)
            logger.debug("[client] Successfully got document.")
            return doc
        except Exception as e:
            logger.exception(
                f"[client] X. Failed to get document by UNID {unid}. Error: {e}")
            return None

    def iter_unids(self, view_names: Optional[List[str]] = None):
        """
        NotesDB から UNID を列挙するジェネレータ
        - 明示ビューは“全部”走査（重複は排除）
        - 明示ビューで1件も取れなければ NoteCollection → それも不可なら AllDocuments
        """
        import traceback

        logger.debug("\n[client] --- Starting iter_unids ---")

        # ★ NameError 対策：必ず先頭で作る
        seen = set()          # 重複UNIDの排除用
        found_any = False     # 何か1件でも見つかったかフラグ

        # ========== A) 明示ビューをすべて走査 ==========
        if view_names:
            logger.debug(f"[client] A. Trying explicit views: {view_names}")
            for vn in view_names:
                try:
                    logger.debug(f"[client] A-1. GetView('{vn}')")
                    view = self.db.GetView(vn)
                except Exception as e:
                    logger.debug(
                        f"[client] A-x. GetView failed for '{vn}': {e}")
                    continue

                if not view:
                    logger.debug(f"[client] A-x. View '{vn}' not found (None)")
                    continue

                try:
                    # View の全エントリ数を確認（環境差で AllEntries が無い場合もある）
                    count = 0
                    try:
                        count = int(view.AllEntries.Count)
                    except Exception:
                        # AllEntries が無い/使えない場合は NthDocument で数えずに走査
                        count = None

                    if count is not None:
                        logger.debug(
                            f"[client] A-2. View '{vn}' has {count} entries.")
                        for i in range(1, count + 1):
                            doc = view.GetNthDocument(i)
                            if doc:
                                uid = doc.UniversalID
                                if uid not in seen:
                                    seen.add(uid)
                                    found_any = True
                                    yield uid
                    else:
                        logger.debug(
                            f"[client] A-2b. View '{vn}' without AllEntries; iterating by GetFirst/Next")
                        doc = view.GetFirstDocument()
                        while doc:
                            uid = doc.UniversalID
                            if uid not in seen:
                                seen.add(uid)
                                found_any = True
                                yield uid
                            doc = view.GetNextDocument(doc)

                    # ★ ここで return しない → 次のビューも続けて見る
                except Exception as e:
                    logger.exception(
                        f"[client] A-x. Iteration error on view '{vn}': {e}")
                    continue

        # ========== B) 明示ビューで1件も取れなかったら NoteCollection ==========
        if not found_any:
            logger.debug(
                "\n[client] B. Explicit views failed or were empty. Trying Fallback 1: CreateNoteCollection")
            try:
                # False で初期化 → 必要な種類だけ選択する
                nc = self.db.CreateNoteCollection(False)

                # --- 環境差異に対応（プロパティ/メソッドの両対応）---
                enabled = False
                # 1) 一般的：データ文書を含める
                try:
                    logger.debug("[client] B-1. Trying 'SelectDocuments = True'")
                    nc.SelectDocuments = True
                    enabled = True
                except Exception as e:
                    logger.debug(
                        f"[client] B-1x. 'SelectDocuments' not available: {e}")

                # 2) メソッド版：SelectAllDataNotes(True)
                if not enabled:
                    try:
                        logger.debug(
                            "[client] B-2. Trying 'SelectAllDataNotes(True)'")
                        nc.SelectAllDataNotes(True)
                        enabled = True
                    except Exception as e:
                        logger.debug(
                            f"[client] B-2x. 'SelectAllDataNotes' not available: {e}")

                # 3) さらに別型：SelectAllNotes(True)
                if not enabled:
                    try:
                        logger.debug("[client] B-3. Trying 'SelectAllNotes(True)'")
                        nc.SelectAllNotes(True)
                        enabled = True
                    except Exception as e:
                        logger.debug(
                            f"[client] B-3x. 'SelectAllNotes' not available: {e}")

                if not enabled:
                    raise RuntimeError(
                        "No suitable selector on INoteCollection (SelectDocuments/SelectAllDataNotes/SelectAllNotes)")

                logger.debug("[client] B-4. Building NoteCollection...")
                nc.BuildCollection()

                count = int(nc.Count)
                logger.debug(
                    f"[client] B-5. NoteCollection built. Found {count} documents.")

                if count > 0:
                    nid = nc.GetFirstNoteId()
                    while nid:
                        doc = self.db.GetDocumentByID(nid)
                        if doc:
                            uid = doc.UniversalID
                            if uid not in seen:   # ← ★ここで NameError が出ていた
                                seen.add(uid)
                                found_any = True
                                yield uid
                        nid = nc.GetNextNoteId(nid)

                    logger.debug("[client] B-6. Finished iterating NoteCollection.")
                    return

            except Exception as e:
                logger.exception(
                    f"[client] X. Fallback 1 (NoteCollection) failed. Error: {e}")

        # ========== C) 最終フォールバック：AllDocuments ==========
        if not found_any:
            logger.debug(
                "[client] C. Fallback 1 failed. Trying Fallback 2: db.AllDocuments")
            try:
                vc = self.db.AllDocuments
                count = int(vc.Count)
                logger.debug(
                    f"[client] C-1. db.AllDocuments collection obtained. Found {count} documents.")
                if count > 0:
                    doc = vc.GetFirstDocument()
                    while doc:
                        uid = doc.UniversalID
                        if uid not in seen:
                            seen.add(uid)
                            yield uid
                        doc = vc.GetNextDocument(doc)
                logger.debug(
                    "[client] C-2. Finished iterating db.AllDocuments. Returning.")
            except Exception as e:
                logger.exception(f"[client] X. Fallback 2 failed. Error: {e}")

        logger.debug("[client] --- Finished iter_unids ---")
