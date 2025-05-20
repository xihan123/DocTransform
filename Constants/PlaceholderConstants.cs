namespace DocTransform.Constants;

public static class PlaceholderConstants
{
    public const string IdCardGender = "{身份证性别}";
    public const string IdCardBirthDate = "{身份证出生日期}";
    public const string IdCardRegion = "{身份证籍贯}";
    public const string IdCardAge = "{身份证年龄}";
    public const string IdCardBirthYear = "{身份证出生年}";
    public const string IdCardBirthMonth = "{身份证出生月}";
    public const string IdCardBirthDay = "{身份证出生日}";

    public static readonly List<string> AllPlaceholders = new()
    {
        IdCardGender,
        IdCardBirthDate,
        IdCardRegion,
        IdCardAge,
        IdCardBirthYear,
        IdCardBirthMonth,
        IdCardBirthDay
    };
}