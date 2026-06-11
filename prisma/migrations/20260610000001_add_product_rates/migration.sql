-- CreateTable
CREATE TABLE "ProductRate" (
    "product_id"       INTEGER NOT NULL,
    "permit_date"      DATE NOT NULL,
    "rate_per_cb"      DECIMAL(10,2) NOT NULL,
    "aroed_per_bottle" DECIMAL(10,4) NOT NULL DEFAULT 0,
    "created_at"       TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "ProductRate_pkey" PRIMARY KEY ("product_id","permit_date")
);

-- AddForeignKey
ALTER TABLE "ProductRate" ADD CONSTRAINT "ProductRate_product_id_fkey"
    FOREIGN KEY ("product_id") REFERENCES "Product"("id")
    ON DELETE RESTRICT ON UPDATE CASCADE;
