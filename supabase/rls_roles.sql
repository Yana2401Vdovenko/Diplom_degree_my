-- ============================================================
-- RLS-політики для ролей: admin, head, laborant, teacher.
-- Виконайте УВЕСЬ файл у Supabase Dashboard → SQL Editor.
-- Після виконання вийдіть із сайту та увійдіть знову,
-- щоб JWT отримав актуальну роль.
-- ============================================================

-- Нормалізуємо історичні коди ролей і українські назви.
CREATE OR REPLACE FUNCTION public.current_app_role()
RETURNS text
LANGUAGE sql
STABLE
AS $$
  SELECT CASE lower(trim(coalesce(
    auth.jwt() -> 'app_metadata' ->> 'role',
    auth.jwt() -> 'user_metadata' ->> 'role',
    ''
  )))
    WHEN 'admin' THEN 'admin'
    WHEN 'head' THEN 'head'
    WHEN 'завідувач кафедри' THEN 'head'
    WHEN 'lab' THEN 'laborant'
    WHEN 'laborant' THEN 'laborant'
    WHEN 'лаборант' THEN 'laborant'
    WHEN 'teacher' THEN 'teacher'
    WHEN 'викладач' THEN 'teacher'
    ELSE ''
  END
$$;

CREATE OR REPLACE FUNCTION public.has_app_role(allowed_roles text[])
RETURNS boolean
LANGUAGE sql
STABLE
AS $$ SELECT public.current_app_role() = any(allowed_roles) $$;

-- RLS не замінює звичайні права PostgreSQL.
GRANT USAGE ON SCHEMA public TO authenticated;
GRANT SELECT ON TABLE public."Посада", public."Тип_викладача", public."Статус_викладача", public."Навантаження" TO authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public."Викладач", public."Кафедральна_інформація_про_виклад", public."Архів" TO authenticated;

-- Вмикаємо RLS лише на таблицях, які реально існують у схемі.
DO $$
DECLARE
  tbl_name text;
  managed_tables text[] := ARRAY[
    'Архів', 'Навантаження', 'Кафедральна_інформація', 'Кафедральна_інформація_про_виклад',
    'Завантаження_навантаження', 'Розподіл_навчального_навантаження',
    'Дисципліни', 'Посада', 'Статус_викладача', 'Тип_викладача', 'Форма_навчання',
    'Рівень_навчання', 'Завантажені_документи', 'Кафедра', 'Викладач',
    'Факультет_ННІ', 'Абревіатура', 'ОПП', 'Спеціальність'
  ];
  documents_tables text[] := ARRAY['Завантажені_документи'];
  laborant_write_tables text[] := ARRAY[
    'Архів', 'Навантаження', 'Кафедральна_інформація', 'Кафедральна_інформація_про_виклад',
    'Завантаження_навантаження', 'Розподіл_навчального_навантаження',
    'Дисципліни', 'Завантажені_документи', 'Кафедра', 'Викладач'
  ];
BEGIN
  FOREACH tbl_name IN ARRAY managed_tables LOOP
    IF to_regclass(format('public.%I', tbl_name)) IS NOT NULL THEN
      EXECUTE format('ALTER TABLE public.%I ENABLE ROW LEVEL SECURITY', tbl_name);
      EXECUTE format('DROP POLICY IF EXISTS app_read ON public.%I', tbl_name);
      EXECUTE format('DROP POLICY IF EXISTS app_write ON public.%I', tbl_name);

      IF tbl_name = ANY(documents_tables) THEN
        -- Документи: читають усі ролі, пишуть admin, laborant і teacher.
        EXECUTE format(
          'CREATE POLICY app_read ON public.%I FOR SELECT TO authenticated
           USING (public.has_app_role(ARRAY[''admin'',''head'',''laborant'',''teacher'']))',
          tbl_name
        );
        EXECUTE format(
          'CREATE POLICY app_write ON public.%I FOR ALL TO authenticated
           USING (public.has_app_role(ARRAY[''admin'',''laborant'',''teacher'']))
           WITH CHECK (public.has_app_role(ARRAY[''admin'',''laborant'',''teacher'']))',
          tbl_name
        );
      ELSIF tbl_name IN ('Кафедральна_інформація', 'Кафедральна_інформація_про_виклад') THEN
        -- Кафедральну інформацію читають усі ролі, редагують admin і laborant.
        EXECUTE format(
          'CREATE POLICY app_read ON public.%I FOR SELECT TO authenticated
           USING (public.has_app_role(ARRAY[''admin'',''head'',''laborant'',''teacher'']))',
          tbl_name
        );
        EXECUTE format(
          'CREATE POLICY app_write ON public.%I FOR ALL TO authenticated
           USING (public.has_app_role(ARRAY[''admin'',''laborant'']))
           WITH CHECK (public.has_app_role(ARRAY[''admin'',''laborant'']))',
          tbl_name
        );
      ELSE
        -- Інші таблиці читають admin, head і laborant.
        EXECUTE format(
          'CREATE POLICY app_read ON public.%I FOR SELECT TO authenticated
           USING (public.has_app_role(ARRAY[''admin'',''head'',''laborant'']))',
          tbl_name
        );
        IF tbl_name = ANY(laborant_write_tables) THEN
          EXECUTE format(
            'CREATE POLICY app_write ON public.%I FOR ALL TO authenticated
             USING (public.has_app_role(ARRAY[''admin'',''laborant'']))
             WITH CHECK (public.has_app_role(ARRAY[''admin'',''laborant'']))',
            tbl_name
          );
        ELSE
          -- Довідники редагує лише admin.
          EXECUTE format(
            'CREATE POLICY app_write ON public.%I FOR ALL TO authenticated
             USING (public.has_app_role(ARRAY[''admin'']))
             WITH CHECK (public.has_app_role(ARRAY[''admin'']))',
            tbl_name
          );
        END IF;
      END IF;
    END IF;
  END LOOP;
END $$;

-- Публічне читання довідників, потрібне формі реєстрації.
DO $$
DECLARE
  pub_tbl text;
  public_read_tables text[] := ARRAY[
    'Кафедра', 'Дисципліни', 'Факультет_ННІ', 'Посада', 'Статус_викладача',
    'Тип_викладача', 'Форма_навчання', 'Рівень_навчання', 'Абревіатура', 'ОПП', 'Спеціальність'
  ];
BEGIN
  FOREACH pub_tbl IN ARRAY public_read_tables LOOP
    IF to_regclass(format('public.%I', pub_tbl)) IS NOT NULL THEN
      EXECUTE format('DROP POLICY IF EXISTS public_read ON public.%I', pub_tbl);
      EXECUTE format(
        'CREATE POLICY public_read ON public.%I FOR SELECT USING (true)',
        pub_tbl
      );
    END IF;
  END LOOP;
END $$;
