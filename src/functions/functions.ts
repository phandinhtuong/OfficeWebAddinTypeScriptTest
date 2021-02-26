/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
/**
 * Returns the input string.
 * @customfunction
 * @param message String to be returned.
 * @returns Input string.
 */
export function show(message: string): string {
  return message;
}
/**
 * Returns the input number plus 15.
 * @customfunction
 * @param number The input number.
 * @returns The input number plus 15.
 */
export function add15(x: number): number{
  return x + 15; 
}
/**
 * Returns author's name and ID in 4 cells.
 * @customfunction
 * @param nameAndID The author's name and ID as a string.
 * @returns The name and ID of author separated.
 */
export function TacGia(nameAndID: string): string[][] {
  if (Number(nameAndID)) //the input is only number
    return [["ERROR: input only number"]];
  
  var splitted = nameAndID.split(" "); //split input by spaces

  if (splitted.length<=1) // input length is not enough
    return [["ERROR: input length <=1"]];

  var authorID = splitted.pop(); // pop the last element / author's ID

  if (isNaN(Number(authorID))) //author's ID is not number
    return [["ERROR: input wrong ID (not number)"]];
  
  var authorName = splitted; //get fullname array

  if (authorName.length==1) //input only last name
    return [["ERROR: only one in name, please input fullname"]];
  
  for (let entry of authorName){//check if name array contains number
    if (!isNaN(Number(entry))) return [["ERROR: number inside name"]];
  }

  var lastNameArray = authorName.pop(); //get last name array
  var lastName=lastNameArray.toString(); //last name in string type

  if (authorName.length==1){ //the remain name is one left / only first name, no middle name
    var firstName = authorName[0].toString();
    return [[firstName,''],[lastName,authorID]]; //this author has no middle name
  }

  //get first name, middle name
  var firstNameArray = authorName[0]; //get first name in array type
  var firstName = firstNameArray.toString(); //first name in string type
  authorName.shift(); //shift to the left / remove first name in array
  var middleName = authorName.join(" ");// get middle name in string
  return [[firstName,middleName],[lastName,authorID]];
}