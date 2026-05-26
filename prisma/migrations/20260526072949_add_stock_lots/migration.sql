-- CreateEnum
CREATE TYPE "StockLotSource" AS ENUM ('OPENING_BALANCE', 'PERMIT');

-- AlterTable
ALTER TABLE "DailySalesBillItem" ADD COLUMN     "stock_lot_id" INTEGER;

-- CreateTable
CREATE TABLE "StockLot" (
    "id" SERIAL NOT NULL,
    "shop_id" INTEGER NOT NULL,
    "product_id" INTEGER NOT NULL,
    "mrp_per_bottle" DECIMAL(10,2) NOT NULL,
    "initial_qty" INTEGER NOT NULL,
    "consumed_qty" INTEGER NOT NULL DEFAULT 0,
    "lot_date" DATE NOT NULL,
    "source" "StockLotSource" NOT NULL,
    "permit_item_id" INTEGER,

    CONSTRAINT "StockLot_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "StockLot_shop_id_product_id_lot_date_idx" ON "StockLot"("shop_id", "product_id", "lot_date");

-- AddForeignKey
ALTER TABLE "DailySalesBillItem" ADD CONSTRAINT "DailySalesBillItem_stock_lot_id_fkey" FOREIGN KEY ("stock_lot_id") REFERENCES "StockLot"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StockLot" ADD CONSTRAINT "StockLot_shop_id_fkey" FOREIGN KEY ("shop_id") REFERENCES "Shop"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StockLot" ADD CONSTRAINT "StockLot_product_id_fkey" FOREIGN KEY ("product_id") REFERENCES "Product"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StockLot" ADD CONSTRAINT "StockLot_permit_item_id_fkey" FOREIGN KEY ("permit_item_id") REFERENCES "PermitItem"("id") ON DELETE SET NULL ON UPDATE CASCADE;
