import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

public class main {
    public static void main(String[] args) {
        // Create a Sikuli screen object
        Screen screen = new Screen();

        try {
            // Open Excel application
            Runtime.getRuntime().exec("cmd /c start excel.exe src/main/resources/TaskData.xlsx");
            Thread.sleep(5000); // Wait for Excel to open

            // Define image patterns for UI elements
            Pattern sortButton = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\sortandfilter.png");
            Pattern customSort = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\customsort.png");
            Pattern sortBy = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\sortBy.png");
            Pattern sortByJoinDateOption = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\joinDate.png");
            Pattern dataOption = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\dataOption.png");
            Pattern removeDuplicatesButton = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\removeDuplicate.png");
            Pattern unSelectAll = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\unSelectAll.png");
            Pattern columnNameCheckBox = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\nameCheckBox.png");
            Pattern okButton = new Pattern("C:\\Users\\Dell\\Desktop\\SecondTaskVois\\images\\okButton.png");

            // Click on the "Sort" button
            screen.click(sortButton);
            screen.click(customSort);
            screen.click(sortBy);
            Thread.sleep(2000);

            // Choose "Join Date" in the "Sort By" dropdown
            screen.click(sortByJoinDateOption);

            // Click on "OK" button to apply the sort
            screen.click(okButton);
            Thread.sleep(2000);

            // Click on "Remove Duplicates" button in the "Data" tab
            screen.click(dataOption);
            screen.click(removeDuplicatesButton);
            Thread.sleep(2000);

            // Check the checkbox for the "Name" column
            screen.click(unSelectAll);
            screen.doubleClick(columnNameCheckBox);

            // Click "OK" to remove duplicates
            screen.click(okButton);

            System.out.println("Excel automation completed successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
