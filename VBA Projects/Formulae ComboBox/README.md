One work related objective is to calculate, on the fly, contract agreed weekly values dividing by 7 and then multiplying by the invoiced number of days. Therefore a textbox is hard coded to calculate the number of days from From and To Date textboxes. A textbox is used to manually enter the figures and formula and it's onAfterUpdate event popluates a results textbox using the Eval function in VBA.

Whilst on Annual Leave I developed a VBA function to use the Eval function in Microsoft Access to calculate text entered in to a textbox in response to a LinkedIn request that improved on my Work related code. Unfortunately the requestor wasn't interested in using it, but it had helped to give me a good insight. Later on I came accross Alessandro's Ms Access image drag videos and I knew that the Classes and code he developed could be modified to enhance my code's process.

ComboBox Text4 is where you enter the formula and figures. By holding the Shift Key and Left Mouse Button down on any of the 3 controls above, then dragging to ComboBox Text4 and releasing the mouse, the name of those controls will be added to the last postion within Text4 you typed. Hence if you already have 1+2 and you add a minus after 1 (1-+2), the control name will be inserted making it 1-ControlName+2. By using he drag and drop method / Control Names, calculations will be more dynamic. The Text4 formula can be saved by either:
 - double clicking will add the current formula if it does not already exist to the rowsource
 - the keys Ctrl D will delete the current formula from the rowsource if it exists
 - both these actions are at runtime, hence the form will switch to design mode, save and close, then re-open

When trying to use the Eval function it should be noted that it does not recognise Form Controls:
 - the code overcomes this by trying to convert them from the error messages in to values before passing back to the Eval function
 - the code allows you to ignore the error messages whilst building the formula and also works when the correct syntax is used for other Ms Access Functions
 - as ListBoxes and ComboBoxes can be multi columned, the code looks for the Listbox.Column(#) patterns and converts them to values before passing to Eval 
 - If no Column(#) is specified with a ListbBox name entry then the first column value of the selected row is used

References
Alessandro Grimaldi
https://www.linkedin.com/feed/update/urn:li:activity:6757210400145608704
