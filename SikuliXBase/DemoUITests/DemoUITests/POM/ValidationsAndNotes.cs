using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoUITests.POM
{
    public class ValidationsAndNotes
    {
        public static string RegistrationCompletedNote = "Your registration completed";
        public static string AddedToCartNote = "The product has been added to your shopping cart";
        public static string AddedToWishlistNote = "The product has been added to your wishlist";
        public static string EmptyWishlistNote = "The wishlist is empty!";
        public static string SearchTermValidationNote = "Search term minimum length is 3 characters";
        public static string NoProductsFoundNote = "No products were found that matched your criteria.";
        public static string DeleteAddressConfirmation = "^Are you sure[\\s\\S]$";
        public static string NoAddressesNote = "No addresses";
    }
}
