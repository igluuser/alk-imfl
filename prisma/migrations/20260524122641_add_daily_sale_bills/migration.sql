-- CreateTable
CREATE TABLE "DailySalesBill" (
    "id" SERIAL NOT NULL,
    "shop_id" INTEGER NOT NULL,
    "bill_date" DATE NOT NULL,
    "bill_year" INTEGER NOT NULL,
    "bill_type" "ProductType" NOT NULL,
    "sequence_no" INTEGER NOT NULL,
    "bill_number" TEXT NOT NULL,
    "total_ml" INTEGER NOT NULL,
    "total_amount" DECIMAL(12,2) NOT NULL,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "DailySalesBill_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DailySalesBillItem" (
    "id" SERIAL NOT NULL,
    "bill_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "bottles" INTEGER NOT NULL,
    "ml_per_bottle" INTEGER NOT NULL,
    "total_ml" INTEGER NOT NULL,
    "mrp_per_bottle" DECIMAL(10,2) NOT NULL,
    "total_amount" DECIMAL(12,2) NOT NULL,

    CONSTRAINT "DailySalesBillItem_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "DailySalesBill_bill_number_key" ON "DailySalesBill"("bill_number");

-- CreateIndex
CREATE INDEX "DailySalesBill_shop_id_bill_date_idx" ON "DailySalesBill"("shop_id", "bill_date");

-- CreateIndex
CREATE UNIQUE INDEX "DailySalesBill_shop_id_bill_year_bill_type_sequence_no_key" ON "DailySalesBill"("shop_id", "bill_year", "bill_type", "sequence_no");

-- AddForeignKey
ALTER TABLE "DailySalesBill" ADD CONSTRAINT "DailySalesBill_shop_id_fkey" FOREIGN KEY ("shop_id") REFERENCES "Shop"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailySalesBillItem" ADD CONSTRAINT "DailySalesBillItem_bill_id_fkey" FOREIGN KEY ("bill_id") REFERENCES "DailySalesBill"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailySalesBillItem" ADD CONSTRAINT "DailySalesBillItem_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
