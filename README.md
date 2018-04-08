# clsSorter
## What is clsSorter used for?
It's a custom class to extend Excel sorting and filtering abilities. Let's see how to use it on an exact example.
## Case description
For example there are three cafes you're managing with. Each of them has the same menu consisting of two types of snacks: sandwiches and cakes. And of course your cafes offer several sorts of coffee-drinks.

Let's imagine that the first cafe is the biggest one, so it's serviced by three employees everyday. For example yesterday Joe, Jane and Jack were on duty. Joe is a barista, Jane makes great sandwiches and Jack is an expert in baking cakes. The second cafe is smaller than the first one. There are only two employees needed there, yesterday it was Kevin and Kate's shift. Kevin is a barista, Kate is responsible for snacks. And the third cafe is the smallest one, so there is only one employee there, it was Mike's duty yesterday. He's a multipurpose specialist.

And here's how the report about yesterday's income looks like:

| id | Cafe | Product | Income |
| :---: | :---: | :---: | :---: |
| 1 | First | Sandwich | 3.25$ |
| 2 | First | Coffee | 2.00$ |
| 3 | Second | Coffee | 2.00$ |
| 4 | Third | Sandwich | 3.50$ |
| 5 | Third | Coffee | 1.50$ |
| 6 | Second | Cake | 3.00$ |
| 7 | First | Cake | 3.00$ |
| 8 | Second | Cake | 3.00$ |
| 9 | Third | Coffee | 2.50$ |
| ... | ... | ... | ... |

You need to find out, which one of your employees gave you the highest income, but as you can see, the problem is that there's no a column containing employee's name. This example is simple enough to create such a column manually. But if the conditions were a little bit more complicated they would make the task rather hard to complete. So let's try to automate this one with help of clsSorter object.

## Decision
### Creating and "feeding"
First of all let's organize the facts we know about our employees in a table:

| Employee's name | Cafe | Product types |
| :---: | :---: | :---: |
| Joe | First | Coffee |
| Jane | First | Sandwich |
| Jack | First | Cake |
| Kevin | Second | Coffee |
| Kate | Second | Sandwich, Cake |
| Mike | Third | Coffee, Sandwich, Cake |

Such a table is a perfect data source for our marker object. Employees' names are obviously should be used as markers and other facts we'll use as conditions. So let's create clsSorter object, and "feed" it with some data:

```
'variables
Dim oSorter As clsSorter
Dim sMarker As String
Dim arrTerms As Variant

'creates a new clsSorter object
Set oSorter = New clsSorter

'adds conditions
sMarker = "Joe"
arrTerms = Array("First", "Coffee")
oSorter.addTerms sMarker, arrTerms
```
Well, it was easy enough. Now clsSorter object "oSorter" knows how to "recognize" Joe basing on information about his workplace and products sold. Of course it's wiser to use a loop to add conditions to "oSorter", but I just want to show you different interesting features of "feeding" clsSorter object. Let's continue with "Kevin" and "Kate" case. You can do it like that:

```
sMarker = "Kevin"
arrTerms = Array("Second", "Coffee")
oSorter.addTerms sMarker, arrTerms

sMarker = "Kate"
arrTerms = Array("Second", "Sandwich, Cake")
oSorter.addTerms sMarker, arrTerms
```
But sometimes it's better to use the exceptation word. Default value of clsSorter object exceptation word is "EXCEPT". Concerning the second cafe we exactly know that Kate is responsible for everything except coffee. So the same results will be achieved with this code:
```
sMarker = "Kevin"
arrTerms = Array("Second", "Coffee")
oSorter.addTerms sMarker, arrTerms

sMarker = "Kate"
arrTerms = Array("Second", "EXCEPT Coffee")
oSorter.addTerms sMarker, arrTerms
```
By the way you can change exceptation word or delimiter (which default value is ",") anytime you want. Just assign any other values to `.delimiter` and `.exceptationWord` properties.

Let's continue with "Mike" case. There are two ways to create conditions, too. Let's see the first one:
```
sMarker = "Mike"
arrTerms = Array("Third", "Coffee, Sandwich, Cake")
oSorter.addTerms sMarker, arrTerms
```
As we know, Mike serves clients alone, so he is resposible for everything in the third cafe. So we just can miss "products" condition. It'll look like:
```
sMarker = "Mike"
arrTerms = Array("Third", Empty)
oSorter.addTerms sMarker, arrTerms
```
Easy, right?
### Marking
All conditions added. What's next? Let's mark each row in our income report. We'll need two columns' values to form requests: "Cafe" and "Product". Just add the forth column to your report and fill it with return values of `.getMarker` method. Here's the example for the row with id number 6:
```
arrTerms = Array("Second", "Cake")
sMarker = oSorter.getMarker(arrTerms)
```
This request will return "Kate" as a marker.
Keep going this way and you'll get needed results:

| id | Cafe | Product | Income | Emplyee's name |
| :---: | :---: | :---: | :---: | :---: |
| 1 | First | Sandwich | 3.25$ | Jane |
| 2 | First | Coffee | 2.00$ | Joe |
| 3 | Second | Coffee | 2.00$ | Kevin |
| 4 | Third | Sandwich | 3.50$ | Mike |
| 5 | Third | Coffee | 1.50$ | Mike |
| 6 | Second | Cake | 3.00$ | Kate |
| 7 | First | Cake | 3.00$ | Jack |
| 8 | Second | Cake | 3.00$ | Kate |
| 9 | Third | Coffee | 2.50$ | Mike |
| ... | ... | ... | ... | ... |

## Summary
That's all for now. Thanks for reading!
