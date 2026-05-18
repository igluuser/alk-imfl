-- CreateEnum
CREATE TYPE "ProductType" AS ENUM ('IMFL', 'BEER');

-- CreateEnum
CREATE TYPE "ImflCategory" AS ENUM ('WHISKY', 'RUM', 'GIN', 'VODKA', 'WINE', 'BRANDY', 'OTHERS');

-- CreateEnum
CREATE TYPE "BeerPackaging" AS ENUM ('BOTTLE', 'TIN');

-- CreateEnum
CREATE TYPE "PermitStatus" AS ENUM ('INDENTED', 'CONFIRMED', 'RECEIVED', 'PARTIAL');

-- CreateEnum
CREATE TYPE "ExpenseCategory" AS ENUM ('TEA', 'PERMIT_RENT', 'POOJA', 'BREAKAGE', 'OVER_CASH', 'ELECTRICITY_BILL', 'GOOGLE_PAY', 'BHATTA', 'GOOGLE_PAY_NEGATIVE', 'PAK', 'OTHERS');

-- CreateEnum
CREATE TYPE "CashCategory" AS ENUM ('COLLECTION', 'DRAWER_CASH', 'TIPS_CASH');

-- CreateEnum
CREATE TYPE "DenominationType" AS ENUM ('NOTE', 'COIN');

-- CreateTable
CREATE TABLE "Shop" (
    "id" SERIAL NOT NULL,
    "name" TEXT NOT NULL,
    "short_name" TEXT NOT NULL,
    "address" TEXT NOT NULL,
    "ksbcl_retailer_id" TEXT NOT NULL,
    "owner_name" TEXT NOT NULL,
    "gst_number" TEXT,
    "license_number" TEXT,
    "is_active" BOOLEAN NOT NULL DEFAULT true,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Shop_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ProductCategory" (
    "id" SERIAL NOT NULL,
    "product_type" "ProductType" NOT NULL,
    "imfl_category" "ImflCategory",
    "beer_packaging" "BeerPackaging",
    "display_name" TEXT NOT NULL,

    CONSTRAINT "ProductCategory_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Product" (
    "id" SERIAL NOT NULL,
    "master_code" INTEGER NOT NULL,
    "short_name" TEXT NOT NULL,
    "ksbcl_item_name" TEXT,
    "ksbcl_item_code" TEXT,
    "category_id" INTEGER NOT NULL,
    "volume_ml" INTEGER NOT NULL,
    "carton_size" INTEGER NOT NULL,
    "aroed_per_bottle" DECIMAL(10,2) NOT NULL DEFAULT 0,
    "is_active" BOOLEAN NOT NULL DEFAULT true,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Product_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Permit" (
    "id" SERIAL NOT NULL,
    "shop_id" INTEGER NOT NULL,
    "indent_no" TEXT NOT NULL,
    "invoice_no" TEXT NOT NULL,
    "permit_date" DATE NOT NULL,
    "bank_transfer_ref" TEXT,
    "bank_transfer_date" DATE,
    "bank_transfer_amount" DECIMAL(12,2),
    "total_indent_cbs" INTEGER NOT NULL DEFAULT 0,
    "total_indent_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "total_cnf_cbs" INTEGER NOT NULL DEFAULT 0,
    "total_cnf_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "has_rationed_items" BOOLEAN NOT NULL DEFAULT false,
    "status" "PermitStatus" NOT NULL DEFAULT 'INDENTED',
    "received_date" DATE,
    "notes" TEXT,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Permit_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PermitItem" (
    "id" SERIAL NOT NULL,
    "permit_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "rate_per_cb" DECIMAL(10,2) NOT NULL,
    "carton_size_snapshot" INTEGER NOT NULL,
    "indent_cbs" INTEGER NOT NULL DEFAULT 0,
    "indent_btls" INTEGER NOT NULL DEFAULT 0,
    "indent_amount" DECIMAL(10,2) NOT NULL,
    "cnf_cbs" INTEGER NOT NULL DEFAULT 0,
    "cnf_btls" INTEGER NOT NULL DEFAULT 0,
    "cnf_amount" DECIMAL(10,2) NOT NULL,
    "is_rationed" BOOLEAN NOT NULL DEFAULT false,
    "mrp_per_bottle" DECIMAL(10,2) NOT NULL,
    "received_btls" INTEGER,

    CONSTRAINT "PermitItem_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Counter" (
    "id" SERIAL NOT NULL,
    "shop_id" INTEGER NOT NULL,
    "name" TEXT NOT NULL,
    "display_order" INTEGER NOT NULL,
    "is_active" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "Counter_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CounterStaffAssignment" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "staff_name" TEXT NOT NULL,
    "assigned_date" DATE NOT NULL,
    "relieved_date" DATE,

    CONSTRAINT "CounterStaffAssignment_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CounterSellingPrice" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "selling_price" DECIMAL(10,2) NOT NULL,
    "effective_from" DATE NOT NULL,
    "effective_to" DATE,

    CONSTRAINT "CounterSellingPrice_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "GodownStock" (
    "id" SERIAL NOT NULL,
    "shop_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "stock_date" DATE NOT NULL,
    "opening_balance_btls" INTEGER NOT NULL DEFAULT 0,
    "received_btls" INTEGER NOT NULL DEFAULT 0,
    "closing_balance_btls" INTEGER NOT NULL DEFAULT 0,

    CONSTRAINT "GodownStock_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "GodownDistribution" (
    "id" SERIAL NOT NULL,
    "godown_stock_id" INTEGER NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "permit_item_id" INTEGER,
    "distributed_btls" INTEGER NOT NULL,
    "distribution_date" DATE NOT NULL,

    CONSTRAINT "GodownDistribution_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DailyCounterStock" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "stock_date" DATE NOT NULL,
    "opening_balance_btls" INTEGER NOT NULL DEFAULT 0,
    "received_from_godown_btls" INTEGER NOT NULL DEFAULT 0,
    "total_btls" INTEGER NOT NULL DEFAULT 0,
    "sold_btls" INTEGER NOT NULL DEFAULT 0,
    "closing_balance_btls" INTEGER NOT NULL DEFAULT 0,
    "mrp_per_bottle" DECIMAL(10,2) NOT NULL,
    "selling_price_per_bottle" DECIMAL(10,2) NOT NULL,
    "sale_amount_mrp" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "tips_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "total_sale_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,

    CONSTRAINT "DailyCounterStock_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DailyCounterSummary" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "summary_date" DATE NOT NULL,
    "staff_name" TEXT,
    "liquor_sale_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "beer_sale_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "total_tips" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "grand_total_by_sale" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "google_pay_amount" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "expenses_total" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "collection_total" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "drawer_cash_total" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "tips_cash_total" DECIMAL(12,2) NOT NULL DEFAULT 0,
    "tally_difference" DECIMAL(12,2) NOT NULL DEFAULT 0,

    CONSTRAINT "DailyCounterSummary_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DailyExpense" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "expense_date" DATE NOT NULL,
    "category" "ExpenseCategory" NOT NULL,
    "amount" DECIMAL(10,2) NOT NULL,
    "is_upi" BOOLEAN NOT NULL DEFAULT false,
    "upi_ref" TEXT,
    "notes" TEXT,

    CONSTRAINT "DailyExpense_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CashRecord" (
    "id" SERIAL NOT NULL,
    "counter_id" INTEGER NOT NULL,
    "record_date" DATE NOT NULL,
    "cash_category" "CashCategory" NOT NULL,
    "denomination_type" "DenominationType" NOT NULL,
    "denomination_value" INTEGER NOT NULL,
    "count" INTEGER NOT NULL,
    "total_amount" INTEGER NOT NULL,

    CONSTRAINT "CashRecord_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "Shop_short_name_key" ON "Shop"("short_name");

-- CreateIndex
CREATE UNIQUE INDEX "Shop_ksbcl_retailer_id_key" ON "Shop"("ksbcl_retailer_id");

-- CreateIndex
CREATE UNIQUE INDEX "Product_master_code_key" ON "Product"("master_code");

-- CreateIndex
CREATE UNIQUE INDEX "Permit_indent_no_key" ON "Permit"("indent_no");

-- CreateIndex
CREATE UNIQUE INDEX "Permit_invoice_no_key" ON "Permit"("invoice_no");

-- CreateIndex
CREATE UNIQUE INDEX "Counter_shop_id_display_order_key" ON "Counter"("shop_id", "display_order");

-- CreateIndex
CREATE INDEX "CounterSellingPrice_counter_id_product_id_effective_from_idx" ON "CounterSellingPrice"("counter_id", "product_id", "effective_from");

-- CreateIndex
CREATE UNIQUE INDEX "GodownStock_shop_id_product_id_stock_date_key" ON "GodownStock"("shop_id", "product_id", "stock_date");

-- CreateIndex
CREATE UNIQUE INDEX "DailyCounterStock_counter_id_product_id_stock_date_key" ON "DailyCounterStock"("counter_id", "product_id", "stock_date");

-- CreateIndex
CREATE UNIQUE INDEX "DailyCounterSummary_counter_id_summary_date_key" ON "DailyCounterSummary"("counter_id", "summary_date");

-- CreateIndex
CREATE INDEX "DailyExpense_counter_id_expense_date_idx" ON "DailyExpense"("counter_id", "expense_date");

-- CreateIndex
CREATE INDEX "CashRecord_counter_id_record_date_idx" ON "CashRecord"("counter_id", "record_date");

-- CreateIndex
CREATE UNIQUE INDEX "CashRecord_counter_id_record_date_cash_category_denominatio_key" ON "CashRecord"("counter_id", "record_date", "cash_category", "denomination_type", "denomination_value");

-- AddForeignKey
ALTER TABLE "Product" ADD CONSTRAINT "Product_category_id_fkey" FOREIGN KEY ("category_id") REFERENCES "ProductCategory"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Permit" ADD CONSTRAINT "Permit_shop_id_fkey" FOREIGN KEY ("shop_id") REFERENCES "Shop"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PermitItem" ADD CONSTRAINT "PermitItem_permit_id_fkey" FOREIGN KEY ("permit_id") REFERENCES "Permit"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "PermitItem" ADD CONSTRAINT "PermitItem_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Counter" ADD CONSTRAINT "Counter_shop_id_fkey" FOREIGN KEY ("shop_id") REFERENCES "Shop"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CounterStaffAssignment" ADD CONSTRAINT "CounterStaffAssignment_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CounterSellingPrice" ADD CONSTRAINT "CounterSellingPrice_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CounterSellingPrice" ADD CONSTRAINT "CounterSellingPrice_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "GodownStock" ADD CONSTRAINT "GodownStock_shop_id_fkey" FOREIGN KEY ("shop_id") REFERENCES "Shop"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "GodownStock" ADD CONSTRAINT "GodownStock_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "GodownDistribution" ADD CONSTRAINT "GodownDistribution_godown_stock_id_fkey" FOREIGN KEY ("godown_stock_id") REFERENCES "GodownStock"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "GodownDistribution" ADD CONSTRAINT "GodownDistribution_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "GodownDistribution" ADD CONSTRAINT "GodownDistribution_permit_item_id_fkey" FOREIGN KEY ("permit_item_id") REFERENCES "PermitItem"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailyCounterStock" ADD CONSTRAINT "DailyCounterStock_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailyCounterStock" ADD CONSTRAINT "DailyCounterStock_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailyCounterSummary" ADD CONSTRAINT "DailyCounterSummary_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DailyExpense" ADD CONSTRAINT "DailyExpense_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CashRecord" ADD CONSTRAINT "CashRecord_counter_id_fkey" FOREIGN KEY ("counter_id") REFERENCES "Counter"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
