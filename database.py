from __future__ import annotations

import os
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterator, Optional


DEFAULT_MONTHLY_QUOTA = 100


@dataclass(frozen=True)
class StudentBalance:
    student_id: int
    base_quota: int
    printed: int
    donated: int
    received: int
    corrected: int
    available: int
    last_print_at: Optional[str]


class Database:
    def __init__(self, db_path: str | os.PathLike | None = None):
        if db_path is None:
            app_dir = Path(os.getenv("LOCALAPPDATA", Path.home())) / "SistemaFotocopias"
            app_dir.mkdir(parents=True, exist_ok=True)
            db_path = app_dir / "sistema_fotocopias.db"
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _initialize(self) -> None:
        with self.connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    last_name TEXT NOT NULL,
                    first_name TEXT NOT NULL,
                    identifier TEXT,
                    active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL
                );

                CREATE INDEX IF NOT EXISTS idx_students_name
                    ON students(last_name, first_name);

                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS movements (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    movement_type TEXT NOT NULL CHECK (
                        movement_type IN ('PRINT', 'DONATION', 'PRINT_REVERSAL')
                    ),
                    student_id INTEGER,
                    related_student_id INTEGER,
                    quantity INTEGER NOT NULL CHECK (quantity > 0),
                    period TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    file_name TEXT,
                    note TEXT,
                    reversed_movement_id INTEGER,
                    operator TEXT,
                    FOREIGN KEY(student_id) REFERENCES students(id),
                    FOREIGN KEY(related_student_id) REFERENCES students(id),
                    FOREIGN KEY(reversed_movement_id) REFERENCES movements(id)
                );

                CREATE UNIQUE INDEX IF NOT EXISTS idx_one_reversal_per_print
                    ON movements(reversed_movement_id)
                    WHERE reversed_movement_id IS NOT NULL;

                CREATE INDEX IF NOT EXISTS idx_movements_period_student
                    ON movements(period, student_id);
                """
            )
            conn.execute(
                "INSERT OR IGNORE INTO settings(key, value) VALUES('monthly_quota', ?)",
                (str(DEFAULT_MONTHLY_QUOTA),),
            )

    @staticmethod
    def normalize_period(year: int, month: int) -> str:
        if month < 1 or month > 12:
            raise ValueError("Mes invalido")
        return f"{year:04d}-{month:02d}"

    def get_monthly_quota(self) -> int:
        with self.connect() as conn:
            row = conn.execute(
                "SELECT value FROM settings WHERE key = 'monthly_quota'"
            ).fetchone()
            return int(row["value"]) if row else DEFAULT_MONTHLY_QUOTA

    def set_monthly_quota(self, quota: int) -> None:
        if quota <= 0:
            raise ValueError("El cupo debe ser mayor que cero")
        with self.connect() as conn:
            conn.execute(
                "INSERT INTO settings(key, value) VALUES('monthly_quota', ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (str(quota),),
            )

    def add_student(self, last_name: str, first_name: str, identifier: str = "") -> int:
        last_name = last_name.strip()
        first_name = first_name.strip()
        identifier = identifier.strip()
        if not last_name or not first_name:
            raise ValueError("Apellido y nombre son obligatorios")
        with self.connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO students(last_name, first_name, identifier, created_at)
                VALUES (?, ?, ?, ?)
                """,
                (last_name, first_name, identifier or None, self._now()),
            )
            return int(cur.lastrowid)

    def deactivate_student(self, student_id: int) -> None:
        with self.connect() as conn:
            conn.execute("UPDATE students SET active = 0 WHERE id = ?", (student_id,))

    def reactivate_student(self, student_id: int) -> None:
        with self.connect() as conn:
            conn.execute("UPDATE students SET active = 1 WHERE id = ?", (student_id,))

    def list_students(self, search: str = "", include_inactive: bool = False):
        params: list[object] = []
        where = []
        if not include_inactive:
            where.append("active = 1")
        if search.strip():
            term = f"%{search.strip()}%"
            where.append(
                "(last_name LIKE ? OR first_name LIKE ? OR COALESCE(identifier, '') LIKE ?)"
            )
            params.extend([term, term, term])
        sql = "SELECT * FROM students"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY last_name COLLATE NOCASE, first_name COLLATE NOCASE"
        with self.connect() as conn:
            return conn.execute(sql, params).fetchall()

    def get_student(self, student_id: int):
        with self.connect() as conn:
            return conn.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()

    def get_balance(self, student_id: int, period: str) -> StudentBalance:
        quota = self.get_monthly_quota()
        with self.connect() as conn:
            printed = conn.execute(
                """
                SELECT COALESCE(SUM(quantity), 0) AS total
                FROM movements
                WHERE movement_type = 'PRINT' AND student_id = ? AND period = ?
                """,
                (student_id, period),
            ).fetchone()["total"]
            corrected = conn.execute(
                """
                SELECT COALESCE(SUM(quantity), 0) AS total
                FROM movements
                WHERE movement_type = 'PRINT_REVERSAL' AND student_id = ? AND period = ?
                """,
                (student_id, period),
            ).fetchone()["total"]
            donated = conn.execute(
                """
                SELECT COALESCE(SUM(quantity), 0) AS total
                FROM movements
                WHERE movement_type = 'DONATION' AND student_id = ? AND period = ?
                """,
                (student_id, period),
            ).fetchone()["total"]
            received = conn.execute(
                """
                SELECT COALESCE(SUM(quantity), 0) AS total
                FROM movements
                WHERE movement_type = 'DONATION' AND related_student_id = ? AND period = ?
                """,
                (student_id, period),
            ).fetchone()["total"]
            row = conn.execute(
                """
                SELECT created_at
                FROM movements
                WHERE movement_type = 'PRINT' AND student_id = ?
                ORDER BY id DESC LIMIT 1
                """,
                (student_id,),
            ).fetchone()
            last_print = row["created_at"] if row else None

        effective_printed = max(0, int(printed) - int(corrected))
        available = quota + int(received) - int(donated) - effective_printed
        return StudentBalance(
            student_id=student_id,
            base_quota=quota,
            printed=effective_printed,
            donated=int(donated),
            received=int(received),
            corrected=int(corrected),
            available=available,
            last_print_at=last_print,
        )

    def donate(
        self,
        donor_id: int,
        recipient_id: int,
        quantity: int,
        period: str,
        note: str = "",
        operator: str = "",
    ) -> int:
        if donor_id == recipient_id:
            raise ValueError("Un becado no puede donarse hojas a si mismo")
        if quantity <= 0:
            raise ValueError("La cantidad debe ser mayor que cero")
        balance = self.get_balance(donor_id, period)
        if balance.available < quantity:
            raise ValueError(
                f"Saldo insuficiente. Disponible: {balance.available} hojas"
            )
        with self.connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO movements(
                    movement_type, student_id, related_student_id, quantity,
                    period, created_at, note, operator
                ) VALUES('DONATION', ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    donor_id,
                    recipient_id,
                    quantity,
                    period,
                    self._now(),
                    note.strip() or None,
                    operator.strip() or None,
                ),
            )
            return int(cur.lastrowid)

    def register_print(
        self,
        student_id: int,
        quantity: int,
        period: str,
        file_name: str = "",
        note: str = "",
        operator: str = "",
    ) -> int:
        if quantity <= 0:
            raise ValueError("La cantidad debe ser mayor que cero")
        balance = self.get_balance(student_id, period)
        if balance.available < quantity:
            raise ValueError(
                f"Saldo insuficiente. Disponible: {balance.available} hojas"
            )
        with self.connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO movements(
                    movement_type, student_id, quantity, period, created_at,
                    file_name, note, operator
                ) VALUES('PRINT', ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    student_id,
                    quantity,
                    period,
                    self._now(),
                    file_name.strip() or None,
                    note.strip() or None,
                    operator.strip() or None,
                ),
            )
            return int(cur.lastrowid)

    def get_last_reversible_print(self, student_id: int, period: str):
        with self.connect() as conn:
            return conn.execute(
                """
                SELECT p.*
                FROM movements p
                LEFT JOIN movements r ON r.reversed_movement_id = p.id
                WHERE p.movement_type = 'PRINT'
                  AND p.student_id = ?
                  AND p.period = ?
                  AND r.id IS NULL
                ORDER BY p.id DESC
                LIMIT 1
                """,
                (student_id, period),
            ).fetchone()

    def reverse_last_print(
        self,
        student_id: int,
        period: str,
        reason: str,
        operator: str = "",
    ) -> int:
        reason = reason.strip()
        if not reason:
            raise ValueError("Debe indicar el motivo de la correccion")
        movement = self.get_last_reversible_print(student_id, period)
        if not movement:
            raise ValueError("No hay impresiones pendientes de correccion en este periodo")
        with self.connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO movements(
                    movement_type, student_id, quantity, period, created_at,
                    file_name, note, reversed_movement_id, operator
                ) VALUES('PRINT_REVERSAL', ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    student_id,
                    movement["quantity"],
                    period,
                    self._now(),
                    movement["file_name"],
                    reason,
                    movement["id"],
                    operator.strip() or None,
                ),
            )
            return int(cur.lastrowid)

    def list_movements(self, period: str, student_id: int | None = None, limit: int = 500):
        sql = """
            SELECT
                m.*,
                s.last_name || ', ' || s.first_name AS student_name,
                rs.last_name || ', ' || rs.first_name AS related_student_name
            FROM movements m
            LEFT JOIN students s ON s.id = m.student_id
            LEFT JOIN students rs ON rs.id = m.related_student_id
            WHERE m.period = ?
        """
        params: list[object] = [period]
        if student_id is not None:
            sql += " AND (m.student_id = ? OR m.related_student_id = ?)"
            params.extend([student_id, student_id])
        sql += " ORDER BY m.id DESC LIMIT ?"
        params.append(limit)
        with self.connect() as conn:
            return conn.execute(sql, params).fetchall()

    @staticmethod
    def _now() -> str:
        return datetime.now().astimezone().isoformat(timespec="seconds")
