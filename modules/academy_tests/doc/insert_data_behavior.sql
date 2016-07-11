BEGIN;

WITH inserted_questions AS (
    INSERT INTO at_question (
        description,
        "name",
        at_topic_id,
        at_level_id,
        create_uid,
        create_date,
        write_uid,
        write_date
    ) SELECT
        description,
        "name",
        at_topic_id,
        at_level_id,
        create_uid,
        create_date,
        write_uid,
        write_date
    FROM non_odoo_at_question RETURNING name, ID
) INSERT into at_category_at_question_rel (at_category_id, at_question_id)
SELECT
    non_odoo_at_question.at_category_id,
    inserted_questions."id" as at_question_id
FROM non_odoo_at_question
INNER JOIN inserted_questions on non_odoo_at_question.name = inserted_questions.name;


INSERT INTO at_answer (
    "name",
    description,
    "sequence",
    is_correct,
    at_question_id,
    create_uid,
    create_date,
    write_uid,
    write_date
) SELECT
    "non_odoo_at_answer"."name",
    "non_odoo_at_answer".description,
    "non_odoo_at_answer"."id" as "sequence",
    "non_odoo_at_answer".is_correct,
    "non_odoo_at_question"."id" as at_question_id,
    "non_odoo_at_answer".create_uid,
    "non_odoo_at_answer".create_date,
    "non_odoo_at_answer".write_uid,
    "non_odoo_at_answer".write_date
FROM non_odoo_at_answer
INNER JOIN non_odoo_at_question on "non_odoo_at_answer".at_question_id = "non_odoo_at_question"."fake_id";

SELECT
    CASE
        WHEN "row_number"() over () = 1 or at_question.name <> LAG(at_question."name") OVER()
        THEN at_question."name" ELSE ''
    END AS question,
    at_answer."name" AS answer
FROM at_question
INNER JOIN at_answer ON at_question."id" = at_answer.at_question_id
WHERE at_question."id" in (select id from non_odoo_at_question)
ORDER BY at_question."id" ASC, at_answer."sequence";

END;
