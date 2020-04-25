package com.company;

public class MyRow {
    String name;
    Integer amount;
    Integer cost;

    public MyRow(String name, double amount, double cost) {
        //if (!name.equals("") && amount != 0 && cost != 0) {
            this.name = name;
            this.amount = (int)Math.round(amount);
            this.cost = (int)Math.round(cost);
        /*} else {
            this.name = "name";
            this.amount = -1;
            this.cost = -1;
        }*/
    }

    public MyRow() {
        this.name = "name";
        this.amount = -1;
        this.cost = -1;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAmount() {
        return amount;
    }

    public void setAmount(Integer amount) {
        this.amount = amount;
    }

    public Integer getCost() {
        return cost;
    }

    public void setCost(Integer cost) {
        this.cost = cost;
    }

    public void printMyRow(){
        System.out.println("name: " + this.name + "\namount: " + this.amount + "\ncost: " + this.cost);
    }
}
