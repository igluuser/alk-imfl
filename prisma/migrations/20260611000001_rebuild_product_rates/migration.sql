-- Drop old combined-PK version
DROP TABLE IF EXISTS "ProductRate";

-- Recreate as change-log: id PK, effective_date is a plain column
CREATE TABLE "ProductRate" (
    "id"               SERIAL PRIMARY KEY,
    "product_id"       INTEGER NOT NULL,
    "effective_date"   DATE NOT NULL,
    "rate_per_cb"      DECIMAL(10,2) NOT NULL,
    "aroed_per_bottle" DECIMAL(10,4) NOT NULL DEFAULT 0,
    "created_at"       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX "ProductRate_product_id_effective_date_idx"
    ON "ProductRate" ("product_id", "effective_date");

ALTER TABLE "ProductRate" ADD CONSTRAINT "ProductRate_product_id_fkey"
    FOREIGN KEY ("product_id") REFERENCES "Product"("id")
    ON DELETE RESTRICT ON UPDATE CASCADE;
