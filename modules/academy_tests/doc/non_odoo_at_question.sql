WITH questions AS (
    SELECT
        t1.name,
        t1.description,
        t1.at_topic_id,
        t1.at_category_id,
        t1.at_level_id
    FROM (
        VALUES
        -------- name, description, topic, category, level ---------
        ('a'::text,'description text'::text,1,1,3),
        ('b'::text,'description text 2'::text,1,2,3)
        ------------------------------------------------------------
    ) t1(name, description, at_topic_id, at_category_id, at_level_id)
) SELECT
    row_number() OVER () AS fake_id,
    questions.name,
    questions.description,
    questions.at_topic_id,
    questions.at_category_id,
    questions.at_level_id,
    1 AS create_uid,
    (to_char(now(), 'YYYY-MM-DD HH:MI:SS.US'::text))::timestamp without time zone AS create_date,
    1 AS write_uid,
    (to_char(now(), 'YYYY-MM-DD HH:MI:SS.US'::text))::timestamp without time zone AS write_date,
    at_question.id
FROM (questions
    LEFT JOIN at_question ON (((at_question.name)::text = questions.name)));
