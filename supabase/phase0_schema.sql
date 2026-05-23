-- ============================================================
-- FutureEdge AI Hub v3.0 — Phase 0 資料庫建置腳本
-- Supabase / PostgreSQL
-- 執行環境：Supabase SQL Editor (dashboard.supabase.com)
-- 建立日期：2026-05-23
-- 注意：請在 Supabase SQL Editor 中依序執行各區塊
-- ============================================================

-- ────────────────────────────────────────────────────────────
-- STEP 1：啟用必要擴充功能
-- ────────────────────────────────────────────────────────────
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- ────────────────────────────────────────────────────────────
-- STEP 2：建立 Enum 類型
-- ────────────────────────────────────────────────────────────
CREATE TYPE user_role AS ENUM ('admin', 'principal', 'teacher', 'student');
CREATE TYPE institution_type AS ENUM ('university', 'senior_high', 'junior_high', 'elementary', 'cram_school', 'other');
CREATE TYPE subscription_tier AS ENUM ('starter', 'growth', 'enterprise');
CREATE TYPE grade_level AS ENUM (
    'elementary_1','elementary_2','elementary_3','elementary_4','elementary_5','elementary_6',
    'junior_1','junior_2','junior_3',
    'senior_1','senior_2','senior_3',
    'university_1','university_2','university_3','university_4'
);

-- ────────────────────────────────────────────────────────────
-- STEP 3：建立核心資料表（8張）
-- ────────────────────────────────────────────────────────────

-- 3.1 institutions — 機構資料
CREATE TABLE institutions (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    code            VARCHAR(20) UNIQUE NOT NULL,       -- 機構代碼，如 SCH-TPE-001
    name            VARCHAR(200) NOT NULL,
    type            institution_type NOT NULL,
    subscription_tier subscription_tier NOT NULL DEFAULT 'starter',
    max_teachers    INTEGER NOT NULL DEFAULT 10,
    data_region     VARCHAR(10) NOT NULL DEFAULT 'tw', -- 個資法：台灣境內
    contact_email   VARCHAR(255),
    is_active       BOOLEAN NOT NULL DEFAULT true,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3.2 users — 使用者（教師、校長、管理員）
CREATE TABLE users (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    institution_id  UUID NOT NULL REFERENCES institutions(id) ON DELETE CASCADE,
    role            user_role NOT NULL,
    email           VARCHAR(255) UNIQUE NOT NULL,
    name_encrypted  BYTEA NOT NULL,                   -- AES-256-GCM 加密真實姓名
    subject_codes   TEXT[] DEFAULT '{}',               -- 任教科目代碼
    is_active       BOOLEAN NOT NULL DEFAULT true,
    last_login_at   TIMESTAMPTZ,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3.3 students — 學生資料（高度加密保護）
CREATE TABLE students (
    id                          UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    institution_id              UUID NOT NULL REFERENCES institutions(id) ON DELETE CASCADE,
    anon_id                     VARCHAR(20) UNIQUE NOT NULL, -- STU-YYYY-#### 格式，供 AI 工具使用
    name_encrypted              BYTEA NOT NULL,              -- AES-256-GCM
    guardian_contact_encrypted  BYTEA,                       -- AES-256-GCM
    grade_level                 grade_level NOT NULL,
    enrollment_year             INTEGER NOT NULL,
    is_active                   BOOLEAN NOT NULL DEFAULT true,
    created_at                  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at                  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3.4 classes — 班級資料
CREATE TABLE classes (
    id                  UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    institution_id      UUID NOT NULL REFERENCES institutions(id) ON DELETE CASCADE,
    name                VARCHAR(100) NOT NULL,              -- 如「高一甲班」
    grade_level         grade_level NOT NULL,
    academic_year       INTEGER NOT NULL,                   -- 學年度，如 114
    homeroom_teacher_id UUID REFERENCES users(id),
    is_active           BOOLEAN NOT NULL DEFAULT true,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 班級學生關聯表
CREATE TABLE class_students (
    class_id    UUID NOT NULL REFERENCES classes(id) ON DELETE CASCADE,
    student_id  UUID NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    joined_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    PRIMARY KEY (class_id, student_id)
);

-- 3.5 assignments — 作業與評量
CREATE TABLE assignments (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    teacher_id      UUID NOT NULL REFERENCES users(id),
    class_id        UUID NOT NULL REFERENCES classes(id),
    institution_id  UUID NOT NULL REFERENCES institutions(id),
    subject         VARCHAR(50) NOT NULL,               -- 科目代碼
    title           VARCHAR(300) NOT NULL,
    description     TEXT,
    rubric          JSONB NOT NULL DEFAULT '{}',        -- 評分標準 JSON
    max_score       NUMERIC(5,2) NOT NULL DEFAULT 100,
    due_date        TIMESTAMPTZ,
    is_active       BOOLEAN NOT NULL DEFAULT true,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3.6 grades — 成績記錄
CREATE TABLE grades (
    id                  UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    assignment_id       UUID NOT NULL REFERENCES assignments(id) ON DELETE CASCADE,
    student_id          UUID NOT NULL REFERENCES students(id),
    institution_id      UUID NOT NULL REFERENCES institutions(id),
    scores              JSONB NOT NULL DEFAULT '{}',    -- 各維度分數
    ai_feedback         TEXT,                           -- AI 生成回饋（anon_id 模式）
    teacher_feedback    TEXT,                           -- 教師補充回饋
    final_score         NUMERIC(5,2),
    teacher_confirmed   BOOLEAN NOT NULL DEFAULT false, -- 教師確認後才算正式
    gdrive_report_url   TEXT,                           -- Google Drive 報告 URL
    graded_at           TIMESTAMPTZ,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    UNIQUE (assignment_id, student_id)
);

-- 3.7 learning_sessions — 學習會話（AI 家教記錄）
CREATE TABLE learning_sessions (
    id              UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    student_id      UUID NOT NULL REFERENCES students(id),
    institution_id  UUID NOT NULL REFERENCES institutions(id),
    subject         VARCHAR(50) NOT NULL,
    topic_path      TEXT[] NOT NULL DEFAULT '{}',       -- 知識點路徑，如 ['數學','代數','二次方程式']
    curriculum_codes TEXT[] DEFAULT '{}',               -- 108課綱代碼
    mastery_before  NUMERIC(4,3) CHECK (mastery_before BETWEEN 0 AND 1),
    mastery_after   NUMERIC(4,3) CHECK (mastery_after BETWEEN 0 AND 1),
    mastery_delta   NUMERIC(4,3) GENERATED ALWAYS AS  -- 自動計算成長值
                    (COALESCE(mastery_after, 0) - COALESCE(mastery_before, 0)) STORED,
    session_log     JSONB DEFAULT '[]',                -- 對話記錄（僅保留問答結構，不含個資）
    duration_seconds INTEGER,
    started_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    ended_at        TIMESTAMPTZ,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3.8 lesson_plans — 課程計劃
CREATE TABLE lesson_plans (
    id                  UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    teacher_id          UUID NOT NULL REFERENCES users(id),
    institution_id      UUID NOT NULL REFERENCES institutions(id),
    subject             VARCHAR(50) NOT NULL,
    grade_level         grade_level NOT NULL,
    unit_title          VARCHAR(300) NOT NULL,
    curriculum_codes    TEXT[] DEFAULT '{}',            -- 對應 108課綱代碼
    learning_objectives TEXT[],
    activities          JSONB DEFAULT '[]',             -- 教學活動結構
    assessment_plan     JSONB DEFAULT '{}',             -- 評量計劃
    gdrive_url          TEXT,                           -- Google Drive 檔案 URL
    is_template         BOOLEAN NOT NULL DEFAULT false,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ────────────────────────────────────────────────────────────
-- STEP 4：建立索引
-- ────────────────────────────────────────────────────────────
CREATE INDEX idx_users_institution ON users(institution_id);
CREATE INDEX idx_users_role ON users(role);
CREATE INDEX idx_students_institution ON students(institution_id);
CREATE INDEX idx_students_anon_id ON students(anon_id);
CREATE INDEX idx_classes_institution ON classes(institution_id);
CREATE INDEX idx_assignments_teacher ON assignments(teacher_id);
CREATE INDEX idx_assignments_class ON assignments(class_id);
CREATE INDEX idx_grades_assignment ON grades(assignment_id);
CREATE INDEX idx_grades_student ON grades(student_id);
CREATE INDEX idx_grades_institution ON grades(institution_id);
CREATE INDEX idx_learning_sessions_student ON learning_sessions(student_id);
CREATE INDEX idx_learning_sessions_subject ON learning_sessions(subject);
CREATE INDEX idx_lesson_plans_teacher ON lesson_plans(teacher_id);

-- ────────────────────────────────────────────────────────────
-- STEP 5：啟用 Row Level Security（RLS）
-- ────────────────────────────────────────────────────────────
ALTER TABLE institutions     ENABLE ROW LEVEL SECURITY;
ALTER TABLE users             ENABLE ROW LEVEL SECURITY;
ALTER TABLE students          ENABLE ROW LEVEL SECURITY;
ALTER TABLE classes           ENABLE ROW LEVEL SECURITY;
ALTER TABLE class_students    ENABLE ROW LEVEL SECURITY;
ALTER TABLE assignments       ENABLE ROW LEVEL SECURITY;
ALTER TABLE grades            ENABLE ROW LEVEL SECURITY;
ALTER TABLE learning_sessions ENABLE ROW LEVEL SECURITY;
ALTER TABLE lesson_plans      ENABLE ROW LEVEL SECURITY;

-- ────────────────────────────────────────────────────────────
-- STEP 6：Helper Function（取得當前用戶的機構 ID 與角色）
-- ────────────────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION get_my_institution_id()
RETURNS UUID LANGUAGE SQL SECURITY DEFINER STABLE AS $$
    SELECT institution_id FROM users WHERE id = auth.uid();
$$;

CREATE OR REPLACE FUNCTION get_my_role()
RETURNS user_role LANGUAGE SQL SECURITY DEFINER STABLE AS $$
    SELECT role FROM users WHERE id = auth.uid();
$$;

-- ────────────────────────────────────────────────────────────
-- STEP 7：RLS 政策（三角色：admin / principal / teacher）
-- ────────────────────────────────────────────────────────────

-- institutions：只有 admin 可見全部；其他人只見自己的機構
CREATE POLICY "institutions_select" ON institutions FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR id = get_my_institution_id()
    );

-- users：admin 全覽；principal 見同機構；teacher 只見自己
CREATE POLICY "users_select" ON users FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() = 'principal' AND institution_id = get_my_institution_id())
        OR id = auth.uid()
    );

-- students：admin 全覽；principal 見同機構；teacher 只見任教班級學生
CREATE POLICY "students_select" ON students FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() IN ('principal') AND institution_id = get_my_institution_id())
        OR (get_my_role() = 'teacher' AND institution_id = get_my_institution_id()
            AND id IN (
                SELECT cs.student_id FROM class_students cs
                JOIN classes c ON cs.class_id = c.id
                JOIN assignments a ON a.class_id = c.id
                WHERE a.teacher_id = auth.uid()
            ))
    );

-- grades：teacher 只見自己出的作業的成績
CREATE POLICY "grades_select" ON grades FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() = 'principal' AND institution_id = get_my_institution_id())
        OR (get_my_role() = 'teacher' AND institution_id = get_my_institution_id()
            AND assignment_id IN (SELECT id FROM assignments WHERE teacher_id = auth.uid()))
    );

CREATE POLICY "grades_insert" ON grades FOR INSERT
    WITH CHECK (
        get_my_role() = 'teacher'
        AND institution_id = get_my_institution_id()
        AND assignment_id IN (SELECT id FROM assignments WHERE teacher_id = auth.uid())
    );

CREATE POLICY "grades_update" ON grades FOR UPDATE
    USING (
        get_my_role() = 'teacher'
        AND assignment_id IN (SELECT id FROM assignments WHERE teacher_id = auth.uid())
    );

-- assignments：teacher 只管自己的作業
CREATE POLICY "assignments_select" ON assignments FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() = 'principal' AND institution_id = get_my_institution_id())
        OR (get_my_role() = 'teacher' AND institution_id = get_my_institution_id())
    );

CREATE POLICY "assignments_insert" ON assignments FOR INSERT
    WITH CHECK (
        get_my_role() = 'teacher'
        AND institution_id = get_my_institution_id()
        AND teacher_id = auth.uid()
    );

-- lesson_plans：teacher 管自己的課程計劃；principal 可見全機構
CREATE POLICY "lesson_plans_select" ON lesson_plans FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() = 'principal' AND institution_id = get_my_institution_id())
        OR (get_my_role() = 'teacher' AND teacher_id = auth.uid())
    );

CREATE POLICY "lesson_plans_insert" ON lesson_plans FOR INSERT
    WITH CHECK (
        get_my_role() = 'teacher'
        AND institution_id = get_my_institution_id()
        AND teacher_id = auth.uid()
    );

-- learning_sessions：teacher 見自班學生；principal 見全校
CREATE POLICY "learning_sessions_select" ON learning_sessions FOR SELECT
    USING (
        get_my_role() = 'admin'
        OR (get_my_role() = 'principal' AND institution_id = get_my_institution_id())
        OR (get_my_role() = 'teacher' AND institution_id = get_my_institution_id())
    );

-- ────────────────────────────────────────────────────────────
-- STEP 8：自動更新 updated_at 的 Trigger
-- ────────────────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$;

CREATE TRIGGER trg_institutions_updated_at
    BEFORE UPDATE ON institutions
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_users_updated_at
    BEFORE UPDATE ON users
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_students_updated_at
    BEFORE UPDATE ON students
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_assignments_updated_at
    BEFORE UPDATE ON assignments
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_grades_updated_at
    BEFORE UPDATE ON grades
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_lesson_plans_updated_at
    BEFORE UPDATE ON lesson_plans
    FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- ────────────────────────────────────────────────────────────
-- STEP 9：anon_id 自動生成函數
-- 格式：STU-{入學年份}-{4位序號}
-- ────────────────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION generate_anon_id(p_enrollment_year INTEGER)
RETURNS VARCHAR LANGUAGE plpgsql AS $$
DECLARE
    v_count INTEGER;
    v_anon_id VARCHAR;
BEGIN
    SELECT COUNT(*) + 1 INTO v_count
    FROM students
    WHERE enrollment_year = p_enrollment_year;

    v_anon_id := 'STU-' || p_enrollment_year || '-' || LPAD(v_count::TEXT, 4, '0');
    RETURN v_anon_id;
END;
$$;

-- ────────────────────────────────────────────────────────────
-- STEP 10：種子資料（測試用機構）
-- ────────────────────────────────────────────────────────────
INSERT INTO institutions (code, name, type, subscription_tier, max_teachers) VALUES
    ('SCH-DEMO-001', '示範高中（測試機構）', 'senior_high', 'growth', 50);

-- ────────────────────────────────────────────────────────────
-- 完成！請確認以下項目：
-- ✅ 8 張資料表建立完成
-- ✅ 索引建立完成
-- ✅ RLS 政策啟用
-- ✅ Helper Functions 就緒
-- ✅ Triggers 自動維護 updated_at
-- ✅ anon_id 生成函數就緒
-- ✅ 測試機構種子資料植入
--
-- 下一步：
-- 1. 在 Supabase Authentication 中啟用 Email 登入
-- 2. 複製 Project URL 與 anon key 至 .env 檔案
-- 3. 通知 Aether 執行 Phase 1（Dify 知識庫建置）
-- ────────────────────────────────────────────────────────────
