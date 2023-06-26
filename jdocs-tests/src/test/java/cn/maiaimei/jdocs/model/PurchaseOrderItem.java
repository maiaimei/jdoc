package cn.maiaimei.jdocs.model;

import java.util.Date;
import lombok.Data;

@Data
public class PurchaseOrderItem {

  public PurchaseOrderItem() {
  }

  public PurchaseOrderItem(Date purchaseTime, String itemNum, String itemName, Double unitPrice,
      Integer count, Double amount, String supplier) {
    this.purchaseTime = purchaseTime;
    this.itemNum = itemNum;
    this.itemName = itemName;
    this.unitPrice = unitPrice;
    this.count = count;
    this.amount = amount;
    this.supplier = supplier;
  }

  private Date purchaseTime;
  private String itemNum;
  private String itemName;
  private Double unitPrice;
  private Integer count;
  private Double amount;
  private String supplier;
}
