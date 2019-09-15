# Factory Planning

Main advantage of being in a big company is having BI tools such as [QlikView](https://www.qlik.com/us/products/qlikview) to extract information about your system. In our case, we're working with [Infor](https://www.infor.com/) as our ERP and get all the data imported to QV every day.

Unfortunately, vast majority of companies don't use this data for more than getting some sales figures. Truth is, you can do much more, specially in operations side.

## Problem to solve

It's really difficult getting an ERP 100% fully functional for your needs (unless you paid a huge amount of money in consulting). Roughly speaking, these type of systems are good for getting an overview of your operations, but not so good on detailed / fine-tunning issues or error spotting.

Planning is one of those fields where everything must be accurate since a slight error at data entry may lead to many mistakes (i.e. changing delivery dates by mistake, mistyping requested product, etc.).

## Solution found

Since we already have a BI system with all the information needed, we created a template to track all this data. New fields have been added for detailed information and colour code have been implemented to check the information more visually.

This system makes the errors easier to spot and allows planning department to check the information at a glance. This is the algorithm flow chart we use:

![Flow chart](https://raw.githubusercontent.com/darroyolpz/Factory-planning/master/20190912_183720.jpg)
