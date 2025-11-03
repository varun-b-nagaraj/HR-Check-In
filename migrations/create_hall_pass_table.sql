CREATE TABLE IF NOT EXISTS hall_passes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    class_id TEXT NOT NULL,
    s_number TEXT NOT NULL,
    name TEXT NOT NULL,
    check_out_time TIMESTAMP NOT NULL,
    check_in_time TIMESTAMP,
    expected_duration INTEGER,  -- in minutes
    actual_duration INTEGER,   -- in minutes, calculated on check-in
    check_out_photo TEXT,
    check_in_photo TEXT,
    check_out_reason TEXT,
    check_in_notes TEXT,
    status TEXT CHECK(status IN ('active', 'completed', 'overdue')) NOT NULL DEFAULT 'active'
);

CREATE INDEX IF NOT EXISTS idx_active_passes ON hall_passes(class_id, s_number, status);
CREATE INDEX IF NOT EXISTS idx_history ON hall_passes(class_id, check_out_time);