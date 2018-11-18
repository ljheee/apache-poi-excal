package com.ljheee.poi.entity;


public class ExcalEntity {

    public long beginDate; // required
    public long endDate; // required
    public String deptName; // required
    public String startAmount; // required
    public String endAmount; // required
    public String procureCost; // required
    public String saleCost; // required
    public String saleAmount; // required
    public String profit; // required
    public String profitRate; // required


    public ExcalEntity(long beginDate, long endDate, String deptName, String startAmount, String endAmount, String procureCost, String saleCost, String saleAmount, String profit, String profitRate) {
        this.beginDate = beginDate;
        this.endDate = endDate;
        this.deptName = deptName;
        this.startAmount = startAmount;
        this.endAmount = endAmount;
        this.procureCost = procureCost;
        this.saleCost = saleCost;
        this.saleAmount = saleAmount;
        this.profit = profit;
        this.profitRate = profitRate;
    }
}
