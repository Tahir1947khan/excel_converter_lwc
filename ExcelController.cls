public class ExcelController {
    
    @AuraEnabled(cacheable=true)
    public static List<Contact> getContactLists() {
        return [SELECT Id, Name, Email, Phone From Contact WHERE Email != null AND Phone != null LIMIT 10];
    }
    
    @AuraEnabled(cacheable=true)
    public static List<Account> getAccountLists() {
        return [SELECT Id, Name, Type, Industry From Account WHERE Type != null AND Industry != null LIMIT 10];
    }
}