-- PostgreSQL 16+. Generated with tools/export_schema.py.

-- Initial empty database only. Use versioned migrations for subsequent changes.

BEGIN;

CREATE TABLE agent (
	id SERIAL NOT NULL, 
	name VARCHAR(120) NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (name)
);

CREATE TABLE carrier (
	id SERIAL NOT NULL, 
	name VARCHAR(120) NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (name)
);

CREATE TABLE import_batch (
	id SERIAL NOT NULL, 
	filename VARCHAR(255) NOT NULL, 
	sha256 VARCHAR(64) NOT NULL, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	summary JSON NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (sha256)
);

CREATE TABLE port (
	id SERIAL NOT NULL, 
	name VARCHAR(120) NOT NULL, 
	is_origin BOOLEAN NOT NULL, 
	is_destination BOOLEAN NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (name)
);

CREATE TABLE quote (
	id VARCHAR(36) NOT NULL, 
	batch_id INTEGER NOT NULL, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	status VARCHAR(20) NOT NULL, 
	mode VARCHAR(20) NOT NULL, 
	snapshot JSON NOT NULL, 
	PRIMARY KEY (id), 
	CONSTRAINT ck_quote_status CHECK (status IN ('draft', 'saved')), 
	CONSTRAINT ck_quote_mode CHECK (mode IN ('strict', 'legacy')), 
	FOREIGN KEY(batch_id) REFERENCES import_batch (id)
);

CREATE TABLE source_row (
	id SERIAL NOT NULL, 
	batch_id INTEGER NOT NULL, 
	sheet VARCHAR(80) NOT NULL, 
	row_number INTEGER NOT NULL, 
	payload JSON NOT NULL, 
	issues JSON NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (batch_id, sheet, row_number), 
	FOREIGN KEY(batch_id) REFERENCES import_batch (id)
);

CREATE INDEX ix_source_row_batch_id ON source_row (batch_id);

CREATE TABLE audit_event (
	id SERIAL NOT NULL, 
	quote_id VARCHAR(36) NOT NULL, 
	created_at TIMESTAMP WITH TIME ZONE NOT NULL, 
	action VARCHAR(40) NOT NULL, 
	detail JSON NOT NULL, 
	PRIMARY KEY (id), 
	FOREIGN KEY(quote_id) REFERENCES quote (id)
);

CREATE INDEX ix_audit_event_quote_id ON audit_event (quote_id);

CREATE TABLE destination_charge (
	id SERIAL NOT NULL, 
	source_row_id INTEGER NOT NULL, 
	batch_id INTEGER NOT NULL, 
	agent_id INTEGER NOT NULL, 
	destination_id INTEGER NOT NULL, 
	carrier_scope VARCHAR(120) NOT NULL, 
	concept VARCHAR(255) NOT NULL, 
	basis VARCHAR(40) NOT NULL, 
	amount NUMERIC(18, 4), 
	status VARCHAR(40) NOT NULL, 
	valid_until DATE, 
	PRIMARY KEY (id), 
	UNIQUE (source_row_id), 
	FOREIGN KEY(source_row_id) REFERENCES source_row (id), 
	FOREIGN KEY(batch_id) REFERENCES import_batch (id), 
	FOREIGN KEY(agent_id) REFERENCES agent (id), 
	FOREIGN KEY(destination_id) REFERENCES port (id)
);

CREATE INDEX ix_destination_charge_batch_id ON destination_charge (batch_id);

CREATE TABLE rate (
	id SERIAL NOT NULL, 
	source_row_id INTEGER NOT NULL, 
	batch_id INTEGER NOT NULL, 
	agent_id INTEGER NOT NULL, 
	carrier_id INTEGER, 
	origin_id INTEGER NOT NULL, 
	destination_id INTEGER NOT NULL, 
	valid_from DATE, 
	valid_until DATE, 
	freight20 NUMERIC(18, 4), 
	freight40st NUMERIC(18, 4), 
	freight40hc NUMERIC(18, 4), 
	documentation NUMERIC(18, 4), 
	currency VARCHAR(12), 
	via VARCHAR(200), 
	transit VARCHAR(120), 
	free_days VARCHAR(120), 
	source_file VARCHAR(255), 
	source_status VARCHAR(40), 
	PRIMARY KEY (id), 
	UNIQUE (source_row_id), 
	FOREIGN KEY(source_row_id) REFERENCES source_row (id), 
	FOREIGN KEY(batch_id) REFERENCES import_batch (id), 
	FOREIGN KEY(agent_id) REFERENCES agent (id), 
	FOREIGN KEY(carrier_id) REFERENCES carrier (id), 
	FOREIGN KEY(origin_id) REFERENCES port (id), 
	FOREIGN KEY(destination_id) REFERENCES port (id)
);

CREATE INDEX ix_rate_batch_id ON rate (batch_id);

CREATE INDEX ix_rate_route_batch ON rate (origin_id, destination_id, batch_id);

COMMIT;
