CREATE
OR REPLACE FUNCTION on_create_question_trigger_function () RETURNS TRIGGER AS $BODY$
BEGIN
    INSERT INTO at_question_variant (
        create_uid,
        create_date,
        write_uid,
        write_date,
        at_question_id,
        alternative_wording,
        active
    ) VALUES (
        NEW.create_uid,
        NEW.create_date,
        NEW.write_uid,
        NEW.write_date,
        NEW."id", NEW."name",
        NEW.active
    );
    RETURN NULL;
END; $BODY$ LANGUAGE plpgsql;


DROP TRIGGER IF EXISTS on_create_question_trigger ON at_question  CASCADE;

CREATE TRIGGER on_create_question_trigger AFTER INSERT ON at_question
FOR EACH ROW EXECUTE PROCEDURE on_create_question_trigger_function ();

------------------------------

CREATE
OR REPLACE FUNCTION on_update_question_trigger_function () RETURNS TRIGGER AS $BODY$
DECLARE
    alt_wording VARCHAR;
BEGIN
    SELECT
        alternative_wording
    FROM
        at_question_variant INTO alt_wording
    WHERE
        at_question_id = NEW."id" ;

    IF alt_wording <> NEW."name" THEN
        RAISE NOTICE 'ID: %  NAME: %', NEW."id", NEW."name";
        UPDATE at_question_variant SET alternative_wording = NEW."name" WHERE at_question_id = NEW."id";
    END IF;

    RETURN NULL;
END; $BODY$ LANGUAGE plpgsql;


DROP TRIGGER IF EXISTS on_update_question_trigger ON at_question  CASCADE;

CREATE TRIGGER on_update_question_trigger AFTER UPDATE ON at_question
FOR EACH ROW EXECUTE PROCEDURE on_update_question_trigger_function ();

------------------------------
