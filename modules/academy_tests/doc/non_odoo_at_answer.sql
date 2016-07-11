WITH answers AS (
    SELECT t1.name,
    t1.description,
    t1.is_correct,
    t1.at_question_id
    FROM (
        VALUES
        --------- name, description, is_correct, question ----------
        ('question text'::text,'description text'::text,true,1),
        ('question text 2'::text,'description text 2'::text,false,1)
        ------------------------------------------------------------
    ) t1(name, description, is_correct, at_question_id)
)
SELECT
    row_number() OVER () AS id,
    answers.name,
    answers.description,
    answers.is_correct,
    answers.at_question_id,
    1 AS create_uid,
    (to_char(now(), 'YYYY-MM-DD HH:MI:SS.US'::text))::timestamp without time zone AS create_date,
    1 AS write_uid,
    (to_char(now(), 'YYYY-MM-DD HH:MI:SS.US'::text))::timestamp without time zone AS write_date
FROM answers;
