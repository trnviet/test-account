using AutoMapper;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Omnia.Foundation.Models.Logging;
using Omnia.HR.Core.Logging;
using Omnia.HR.Model.Enums;
using Omnia.HR.Model.Models;
using Omnia.HR.Model.Models.SPSearch;
using Omnia.HR.Repositories.Infrastructures;
using Omnia.HR.Services.Infrastructures;
using Omnia.HR.SPRepositories;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using static Omnia.HR.Core.Constants;
using Constants = Omnia.HR.Core.Constants;
using Entities = Omnia.HR.Repositories.Entities;
using RoleType = Omnia.HR.Model.Enums.RoleType;
using User = Omnia.HR.Model.Models.User;
using UserRole = Omnia.HR.Model.Models.UserRole;

namespace Omnia.HR.Services.Services
{

    public interface IUserService : IBaseService
    {
        User AddUser(ClientContext ctx, UserPost user);
        void UpdateUser(User user);
        User ReactivateUserByLoginName(ClientContext ctx, string loginName);
        User GetUserById(Guid userId, Guid authouringSiteId);
        List<User> SearchSpUser(ClientContext ctx, string searchString);
        User GetUserByLoginName(ClientContext ctx, string loginName);
        User GetUserByLoginNameInAuthoringSite(ClientContext ctx, string loginName, Guid authoringSiteId);
        User GetUserByLoginNameInAuthoringSitePublic(string loginName, Guid authoringSiteId);
        List<UserBaseModel> GetUsersByDelegate(Guid userId);
        List<User> GetAllUserInAuthoringSite(Guid authoringSiteId);
        List<User> GetAllUserInDefaultAuthoringSite(Guid authoringSiteId);
        List<User> GetAllManagerInAuthoringSite(Guid authoringSiteId);
        List<TeamMember> GetTeamMemberByUserId(Guid userId);
        List<UserWithInformationModel> GetAllUsers(Guid tenantId);
        List<UserWithInformationModel> GetAllActiveUsers(Guid tenantId);
        bool CheckUserExistedInTenant(string loginName, Guid tenantId);
        bool CheckUserExistedInAuthoringSite(Guid authoringSiteId, string loginName);
        string EnsureSpLoginName(ClientContext ctx, string loginName);
        void AddUserToSharePointGroup(ClientContext ctx, User user, string userRole, string authoringSiteName);
        void RemoveUserFromSharePointGroup(ClientContext ctx, string loginName, string userRole,
            string authoringSiteName);
        UserWithUserLeaveTrackPaging GetUsersWithUserLeaveTrack(int page, int itemsPerPage, Guid authoringSiteId);
        UserWithUserLeaveTrack GetUserWithUserLeaveTrackByUserIdAndYearId(Guid authoringSiteId, Guid userId, int yearId);
        UserWithUserLeaveTrackPaging GetUserWithUserLeaveTrackByFilter(int page, int itemsPerPage, Guid authoringSiteId, LeavesFilter filter);
        UserWithUserLeaveTrackPaging GetUserWithUserLeaveTrackByFilterForExport(int page, int itemsPerPage, Guid authoringSiteId, LeavesFilter filter);
        UserPaging GetUsersByFilter(int page, int itemsPerPage, Guid siteId, UsersFilter filter);
        List<User> SearchNonSPUserInAuthoringSite(SearchUserInAuthoringSite model);
        List<UserWithInformationModel> SearchUserInAuthoringSite(SearchUserInAuthoringSite model);
        bool UpdateDelegateUser(DelegateUsersPost delegateUsers);
        void RemoveUser(ClientContext ctx, Omnia.HR.Repositories.Entities.User user, string userRole, Guid authoringSiteId, string authoringSiteName);
        UserRole AddUserRole(User newUser, Guid authoringId, int? roleId, bool isAuthoringSite);
        DelegateUser AddDelegateUser(Guid delegateId, Guid delegateForUserId);
        void AddSPUser(DefaultSetting azCred, UserPost user);
        string GetUserEmailFromSharepoint(ClientContext ctx);
        void SendInvitationEmail(UserPost user, DefaultSetting azCred, string language, string siteUrl);
        bool CanCreateSharepointFolder(Guid userId, Guid authoringSiteId);
        void DeactivateUser(Guid userId, Guid authouringSiteId, ClientContext clientContext, string siteName, Guid modifiedByUserId);
        IList<ExtraEmail> GetEmailFromSPUser(ClientContext ctx);
        User GetUserByEmail(ClientContext ctx, string email);
        bool CreateUserFolder(Guid userId, ClientContext ctx, Guid authoringSiteId);
        void UpdateAzureUserIdForMissUsers();
    }

    public class UserService : BaseService, IUserService
    {

        private readonly ILogger _logger;
        private readonly ISPUtilitiesRepository _spUtilitiesRepository;
        private readonly ISPSearchServiceRepository _spSearchServiceRepository;
        private readonly ISPUserRepository _spUserRepository;
        private readonly IUserRoleService _userRoleService;
        private readonly ILeaveService _leaveService;
        public UserService(IUnitOfWork unitOfWork,
            ILogger logger,
            ISPSearchServiceRepository spSearchServiceRepository,
            ISPUserRepository spUserRepository, IUserRoleService userRoleService, ILeaveService leaveService, ISPUtilitiesRepository spUtilitiesRepository)
            : base(unitOfWork)
        {
            _leaveService = leaveService;
            _logger = logger;
            _spSearchServiceRepository = spSearchServiceRepository;
            _spUserRepository = spUserRepository;
            _userRoleService = userRoleService;
            _spUtilitiesRepository = spUtilitiesRepository;
        }

        public User GetUserByEmail(ClientContext ctx, string email)
        {
            try
            {
                var web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                var userSP = web.SiteUsers.GetByEmail(email);
                ctx.Load(userSP);
                ctx.ExecuteQuery();

                var user = _unitOfWork.UserRepository.Find(x => x.LoginName == userSP.LoginName).Include(x => x.UserInformation).FirstOrDefault();
                return Mapper.Map<User>(user);

            }
            catch (Exception ex)
            {
                var userInfomations = _unitOfWork.UserInformationRepository.Find(x => x.Email == email).Select(x => x.Id).ToList();

                var user = _unitOfWork.UserRepository.Find(x => x.UserInformationId.HasValue && userInfomations.Contains(x.UserInformationId.Value)).Include(x => x.UserInformation).FirstOrDefault();
                return Mapper.Map<User>(user);
            }
        }

        public List<UserWithInformationModel> GetAllUsers(Guid tenantId)
        {
            var user = GetAllUsersContent(tenantId).ToList();
            var result = Mapper.Map<List<UserWithInformationModel>>(user);
            return result;
        }

        public List<UserWithInformationModel> GetAllActiveUsers(Guid tenantId)
        {
            var user = GetAllUsersContent(tenantId).Where(t => t.IsActive).ToList();
            var result = Mapper.Map<List<UserWithInformationModel>>(user);
            return result;
        }

        public User AddUser(ClientContext ctx, UserPost user)
        {
            try
            {

                var newUser = AddUserDataMapping(ctx, user);
                var userEntity = Mapper.Map<Entities.User>(newUser);
                userEntity.Created = DateTime.UtcNow;
                userEntity.Modified = DateTime.UtcNow;
                //userEntity.AzureADUserId = GetAzureADUserId(user.Email);
                _unitOfWork.UserRepository.Add(userEntity);
                _unitOfWork.Save();

                return newUser;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddUser", ex.StackTrace);
                throw;
            }
        }

        public User ReactivateUserByLoginName(ClientContext ctx, string loginName)
        {
            try
            {
                var spLoginName = EnsureSpLoginName(ctx, loginName);
                var userEntity = _unitOfWork.UserRepository
                    .Find(t => t.LoginName == spLoginName)
                    .FirstOrDefault();

                if (userEntity != null) userEntity.IsActive = true;
                _unitOfWork.Save();
                return Mapper.Map<User>(userEntity);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "ReactivateUser", ex.StackTrace);
                throw;
            }
        }
        public string GetUserEmailFromSharepoint(ClientContext ctx)
        {
            var web = ctx.Web;
            var user = web.CurrentUser;
            ctx.Load(web);
            ctx.Load(user);
            ctx.ExecuteQuery();
            return user.Email;
        }
        public void DeactivateUser(Guid userId, Guid authouringSiteId, ClientContext clientContext, string siteName, Guid modifiedByUserId)
        {
            try
            {
                // Update user
                var user = _unitOfWork.UserRepository.FindById(userId);
                user.IsActive = false;
                _unitOfWork.UserRepository.Update(user);
                _unitOfWork.Save();
                // deactive contract
                var contracts = _unitOfWork.ContractRepository.Find(x => x.UserId == userId).ToList();
                contracts.ForEach(x => x.Status = ContractStatusEnum.InActive);
                _unitOfWork.Save();
                // deactive payslip
                var payslips = _unitOfWork.PayslipRepository.Find(x => x.UserId == userId).ToList();
                payslips.ForEach(x => x.Status = PayslipStatusEnum.Draft);
                _unitOfWork.Save();
                // reject all leaves send to this user
                var leaves = (from leaveRequestApprover in _unitOfWork.LeaveRequestApproverRepository.DbSet()
                              join leaveRequest in _unitOfWork.LeaveRequestRepository.DbSet()
                                  on leaveRequestApprover.LeaveRequestId equals leaveRequest.Id
                              where leaveRequestApprover.UserId == userId && (leaveRequest.Status == ApprovalStep.Pending && !leaveRequest.IsRemoved)
                              select leaveRequest).ToList();

                leaves.ForEach(x => x.Status = ApprovalStep.Rejected);
                _unitOfWork.Save();
                foreach (var leave in leaves)
                {
                    var leaveModel = Mapper.Map<LeaveRequest>(leave);

                    _leaveService.UpdateUserLeaveTrackBy(leaveModel, UserAction.Reject, authouringSiteId, modifiedByUserId);

                    _leaveService.AddLeaveRequestHistory(new LeaveRequestHistory
                    {
                        Action = UserAction.Reject,
                        LeaveRequestId = leave.Id,
                        UserId = userId,
                        Comment = Constants.LeaveHistory.SystemRejected

                    }, UserAction.Reject);
                }
                // remove all leave requests

                var requests = _unitOfWork.LeaveRequestRepository.Find(x => x.RequestForUserId == userId && x.Status == ApprovalStep.Pending && !x.IsRemoved).ToList();
                requests.ForEach(x => x.IsRemoved = true);
                _unitOfWork.Save();
                foreach (var leave in requests)
                {
                    var leaveModel = Mapper.Map<LeaveRequest>(leave);

                    _leaveService.UpdateUserLeaveTrackBy(leaveModel, UserAction.Remove, authouringSiteId, modifiedByUserId);

                    _leaveService.AddLeaveRequestHistory(new LeaveRequestHistory
                    {
                        Action = UserAction.Remove,
                        UserId = leave.RequestForUserId,
                        LeaveRequestId = leave.Id,
                        Comment = Constants.LeaveHistory.SystemRemove

                    }, UserAction.Remove);
                }


                /// Remove Roles
                var userRoles = _userRoleService.GetUserRolesByUserAuthoringSite(user.Id, authouringSiteId, null);

                if (userRoles.Any())
                {
                    foreach (var userRole in userRoles)
                    {
                        RemoveUser(clientContext, user, userRole.Role.Name,
                                authouringSiteId, siteName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "DeactivateUser", ex.StackTrace);
                throw;
            }

        }

        public User GetUserById(Guid userId, Guid authouringSiteId)
        {
            try
            {
                var result = (from user in _unitOfWork.UserRepository.DbSet()
                              where user.Id == userId
                              select user)
                                            .Include(t => t.UserInformation)
                                            .Include(t => t.Department)
                                            .Include(t => t.JobTitle)
                                            .Include(t => t.OfficeLocation)
                                            .FirstOrDefault();
                var users = Mapper.Map<User>(result);
                return users;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserById", ex.StackTrace);
                throw;
            }
        }

        public List<User> SearchSpUser(ClientContext ctx, string searchString)
        {
            try
            {
                var filter = new UserFilter
                {
                    SearchString = searchString
                };

                var result = _spSearchServiceRepository.SearchSpUser(ctx, filter);
                return result.Users;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "SearchSpUser", ex.StackTrace);
                throw;
            }
        }

        public User GetUserByLoginName(ClientContext ctx, string loginName)
        {
            // loginName should be ensured; domain/user <-> user
            var spLoginName = EnsureSpLoginName(ctx, loginName);
            var result = _unitOfWork.UserRepository
                .Find(t => t.LoginName == spLoginName)
                .FirstOrDefault();
            return Mapper.Map<Model.Models.User>(result);
        }

        public User GetUserByLoginNameInAuthoringSite(ClientContext ctx, string loginName, Guid authoringSiteId)
        {
            // loginName should be ensured; domain/user <-> user
            var spLoginName = EnsureSpLoginName(ctx, loginName);
            User result = spLoginName != null ? GetUserByLoginNameInAuthoringSitePublic(spLoginName, authoringSiteId) : null;
            return result;
        }

        public User GetUserByLoginNameInAuthoringSitePublic(string loginName, Guid authoringSiteId)
        {
            var user = _unitOfWork.UserRepository
                .Find(t => t.LoginName == loginName, includeProperties: "UserRoles")
                .FirstOrDefault();
            var isInAuthoringSite = user?.UserRoles.Where(t => t.AuthoringSiteId == authoringSiteId).FirstOrDefault().IsNotNull();
            if (user == null)
            {
                return null;
            }
            var result = Mapper.Map<Model.Models.User>(user);
            result.IsExistedInAuthoringSite = isInAuthoringSite.GetValueOrDefault();
            return result;
        }

        public List<UserBaseModel> GetUsersByDelegate(Guid userId)
        {
            try
            {
                var usersByDelegateQuery = from delegateUser in _unitOfWork.DelegateUserRepository.DbSet()
                                           join user in _unitOfWork.UserRepository.DbSet()
                                               on delegateUser.DelegateForUserId equals user.Id
                                           where delegateUser.DelegateId == userId && user.IsActive
                                           select user;
                return Mapper.Map<List<UserBaseModel>>(usersByDelegateQuery.OrderByDescending(o => o.Id == userId).ThenBy(o => o.Id).ToList());
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserByDelegate", ex.StackTrace);
                throw;
            }
        }

        public List<User> GetAllUserInAuthoringSite(Guid authoringSiteId)
        {
            try
            {
                var query = (from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet()
                                 on user.Id equals userRole.UserId
                             where userRole.AuthoringSiteId == authoringSiteId
                             select user).Include(t => t.UserInformation);
                var result = Mapper.Map<List<Model.Models.User>>(query.ToList());
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetAllUserInAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public List<User> GetAllUserInDefaultAuthoringSite(Guid authoringSiteId)
        {
            try
            {
                var query = (from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet()
                                 on user.Id equals userRole.UserId
                             where userRole.AuthoringSiteId == authoringSiteId && userRole.IsDefault
                             select user).Include(t => t.UserInformation);
                var result = Mapper.Map<List<Model.Models.User>>(query.ToList());
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetAllUserInDefaultAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public List<User> GetAllManagerInAuthoringSite(Guid authoringSiteId)
        {
            try
            {
                var managerRole = _unitOfWork.RoleRepository.Find(t => t.Name == RoleType.Manager.ToString()).FirstOrDefault();

                var query = (from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet()
                                 on user.Id equals userRole.UserId
                             where userRole.AuthoringSiteId == authoringSiteId && userRole.RoleId == managerRole.Id
                             select user).Include(t => t.UserInformation);

                var result = Mapper.Map<List<User>>(query.ToList());
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetAllManagerInAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public bool CheckUserExistedInTenant(string loginName, Guid tenantId)
        {
            try
            {
                var result = from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet()
                                 on user.Id equals userRole.UserId
                             join authoringSite in _unitOfWork.AuthoringSiteRepository.DbSet()
                                on userRole.AuthoringSiteId equals authoringSite.Id
                             where user.LoginName == loginName
                                   && userRole.IsDefault
                                   && authoringSite.TenantId == tenantId
                             select user;
                return result.FirstOrDefault() != null;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "CheckUserExistedInTenant", ex.StackTrace);
                throw;
            }
        }

        public bool CheckUserExistedInAuthoringSite(Guid authoringSiteId, string loginName)
        {
            try
            {
                var result = from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet()
                                 on user.Id equals userRole.UserId
                             where user.LoginName == loginName
                                   && userRole.AuthoringSiteId == authoringSiteId
                             select user;
                return result.FirstOrDefault() != null;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "CheckUserExistedInAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public string EnsureSpLoginName(ClientContext ctx, string loginName)
        {
            try
            {
                var ensuredLoginName = _spUserRepository.EnsureSpLoginName(ctx, loginName);
                return ensuredLoginName;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void AddUserToSharePointGroup(ClientContext ctx, User user, string userRole, string authoringSiteName)
        {
            try
            {
                _spUserRepository.AddUserToSharePointGroup(ctx, user, userRole, authoringSiteName);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddUserToSharePointGroup", ex.StackTrace);
                throw;
            }
        }

        public void RemoveUserFromSharePointGroup(ClientContext ctx, string loginName, string userRole,
            string authoringSiteName)
        {
            try
            {
                _spUserRepository.RemoveUserFromSharePointGroup(ctx, loginName, userRole, authoringSiteName);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "RemoveUserFromSharePointGroup", ex.StackTrace);
                throw;
            }
        }

        public List<TeamMember> GetTeamMemberByUserId(Guid userId)
        {
            try
            {
                // TODO: CC: Include everything in one query
                var myGroupDetailListEntity = _unitOfWork.GroupDetailRepository
                    .Find(t => t.UserId == userId)
                    .Include("Group.GroupDetails.User.UserInformation")
                    .Include("Group.GroupDetails.LeaveApprovalLevel.LeaveApprovalLevelLocalizations")
                    .ToList();
                var myGroupDetailList = Mapper.Map<List<GroupDetail>>(myGroupDetailListEntity);

                var memberList = new List<TeamMember>();
                foreach (var groupDetail in myGroupDetailList)
                {
                    // mapping level for users in group
                    groupDetail.Group.GroupDetails = groupDetail.Group.GroupDetails.Select(t =>
                    {
                        // default language is 'en-US'
                        t.User.Level = t.LeaveApprovalLevel.LeaveApprovalLevelLocalizations[0].Name;
                        return t;
                    }).ToList();
                    // remove current user
                    var userList = groupDetail.Group.GroupDetails
                        .Where(t => t.User.Id != userId)
                        .Select(t =>
                        {
                            // lighten user object
                            var user = t.User;
                            return user;
                        })
                        .ToList();
                    // team member list
                    var member = new TeamMember
                    {
                        Id = groupDetail.Id,
                        Name = groupDetail.Group.Name,
                        User = userList,
                        IsFavorite = groupDetail.IsFavorite
                    };
                    memberList.Add(member);
                }
                return memberList;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetTeamMemberByUserId", ex.StackTrace);
                throw;
            }
        }

        public UserWithUserLeaveTrackPaging GetUsersWithUserLeaveTrack(int page, int itemsPerPage, Guid authoringSiteId)
        {
            try
            {
                var userPaging = new UserWithUserLeaveTrackPaging();
                var result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet()
                              on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                              select user)
                    .Include(t => t.UserInformation)
                    .Include(t => t.UserLeaveTracks);

                var totalCount = result.Count();
                userPaging.TotalItems = totalCount;
                userPaging.TotalPages = (int)Math.Ceiling((double)totalCount / itemsPerPage);
                var latestYear = GetLatestUserLeaveYear(authoringSiteId);
                if (page <= userPaging.TotalPages && page >= 1)
                {
                    if (totalCount > itemsPerPage)
                    {
                        var userEntities = result?.OrderByDescending(t => t.Id).Skip(itemsPerPage * (page - 1))
                            .Take(itemsPerPage)
                            .ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracks(userEntities, latestYear);
                    }
                    else
                    {
                        var userEntities = result?.ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracks(userEntities, latestYear);
                    }
                }
                return userPaging;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUsersWithUserLeaveTrack", ex.StackTrace);
                throw;
            }
        }

        public UserWithUserLeaveTrack GetUserWithUserLeaveTrackByUserIdAndYearId(Guid authoringSiteId, Guid userId, int yearId)
        {
            try
            {
                var query = (from user in _unitOfWork.UserRepository.DbSet()
                             join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                             where user.Id == userId
                                   && user.IsActive == true
                                   && userRole.AuthoringSiteId == authoringSiteId
                                   && userRole.IsDefault == true
                             select user)
                    .Include(t => t.UserInformation)
                    .Include(t => t.UserLeaveTracks);

                var userEntity = query.FirstOrDefault();
                var result = MapToUserWithUserLeaveTrack(userEntity, yearId);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserWithUserLeaveTrackByUserId", ex.StackTrace);
                throw;
            }
        }

        public UserWithUserLeaveTrackPaging GetUserWithUserLeaveTrackByFilter(int page, int itemsPerPage, Guid authoringSiteId, LeavesFilter filter)
        {
            try
            {
                var userPaging = new UserWithUserLeaveTrackPaging();
                IQueryable<Entities.User> result;
                if (filter.IsManager)
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                              select user)
                        .Include(t => t.UserInformation)
                        .Include(t => t.UserLeaveTracks);
                }
                else
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                                    && user.Id == filter.UserId
                              select user)
                         .Include(t => t.UserInformation)
                         .Include(t => t.UserLeaveTracks);
                }
                //
                if (!string.IsNullOrEmpty(filter.SearchString))
                {
                    result = from user in result
                             where user.DisplayName.Trim().ToLower().Contains(filter.SearchString.ToLower())
                             select user;
                }
                //
                var totalCount = result.Count();
                userPaging.TotalItems = totalCount;
                userPaging.TotalPages = (int)Math.Ceiling((double)totalCount / itemsPerPage);
                // Filter By current year
                var years = new List<UserLeaveYear>();
                if (!filter.StartDate.HasValue && !filter.EndDate.HasValue)
                {
                    years = GetUserLeaveYearByYear(authoringSiteId, filter.Year);
                }
                else
                {
                    years = GetUserLeaveYearByRange(authoringSiteId, filter.StartDate, filter.EndDate);
                }
                if (page <= userPaging.TotalPages && page >= 1)
                {
                    if (totalCount > itemsPerPage)
                    {
                        var userEntities = result?.OrderByDescending(t => t.DisplayName).Skip(itemsPerPage * (page - 1))
                            .Take(itemsPerPage)
                            .ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracks(userEntities, years, filter);
                    }
                    else
                    {
                        var userEntities = result?.ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracks(userEntities, years, filter);
                    }
                }
                return userPaging;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserWithUserLeaveTrackByFilter", ex.StackTrace);
                throw;
            }
        }

        public UserWithUserLeaveTrackPaging GetUserWithUserLeaveTrackByFilterForExport(int page, int itemsPerPage, Guid authoringSiteId, LeavesFilter filter)
        {
            try
            {
                var userPaging = new UserWithUserLeaveTrackPaging();
                IQueryable<Entities.User> result;
                if (filter.IsManager)
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                              select user)
                        .Include(t => t.UserInformation)
                        .Include(t => t.UserLeaveTracks);
                }
                else
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                                    && user.Id == filter.UserId
                              select user)
                         .Include(t => t.UserInformation)
                         .Include(t => t.UserLeaveTracks);
                }
                //
                if (!string.IsNullOrEmpty(filter.SearchString))
                {
                    result = from user in result
                             where user.DisplayName.Trim().ToLower().Contains(filter.SearchString.ToLower())
                             select user;
                }
                //
                var totalCount = result.Count();
                userPaging.TotalItems = totalCount;
                userPaging.TotalPages = (int)Math.Ceiling((double)totalCount / itemsPerPage);
                // Filter By current year
                var years = new List<UserLeaveYear>();
                if (!filter.StartDate.HasValue && !filter.EndDate.HasValue)
                {
                    years = GetUserLeaveYearByYear(authoringSiteId, filter.Year);
                }
                else
                {
                    years = GetUserLeaveYearByRange(authoringSiteId, filter.StartDate, filter.EndDate);
                }
                if (page <= userPaging.TotalPages && page >= 1)
                {
                    if (totalCount > itemsPerPage)
                    {
                        var userEntities = result?.OrderByDescending(t => t.DisplayName).Skip(itemsPerPage * (page - 1))
                            .Take(itemsPerPage)
                            .ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracksForExport(userEntities, years, filter);
                    }
                    else
                    {
                        var userEntities = result?.ToList();
                        userPaging.Users = MapToUserWithUserLeaveTracksForExport(userEntities, years, filter);
                    }
                }
                return userPaging;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserWithUserLeaveTrackByFilter", ex.StackTrace);
                throw;
            }
        }

        public void SendInvitationEmail(UserPost user, DefaultSetting azCred, string language, string siteUrl)
        {
            try
            {
                AzureAuthenticationProvider.SendInvitation(user, siteUrl, language, azCred);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddSPUser", ex.StackTrace);
                throw;
            }
        }

        public UserPaging GetUsersByFilter(int page, int itemsPerPage, Guid siteId, UsersFilter filter)
        {
            try
            {
                var userPaging = new UserPaging();
                IQueryable<Entities.User> result;
                if (siteId != Guid.Empty)
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet()
                              on user.Id equals userRole.UserId
                              where user.IsActive == true && userRole.AuthoringSiteId == siteId && userRole.IsDefault == true
                              select user);
                }
                else
                {
                    result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet()
                              on user.Id equals userRole.UserId
                              where user.IsActive == true && userRole.IsDefault == true
                              select user);
                }
                result = FilterUser(filter, result);
                var totalCount = result.Count();
                userPaging.TotalItems = totalCount;
                userPaging.TotalPages = (int)Math.Ceiling((double)totalCount / itemsPerPage);
                if (page <= userPaging.TotalPages && page >= 1)
                {
                    if (totalCount > itemsPerPage)
                    {
                        var users = result?.OrderByDescending(t => t.StartDate).Skip(itemsPerPage * (page - 1))
                            .Take(itemsPerPage)
                            .Include(t => t.UserInformation).AsNoTracking()
                            .Include(t => t.JobTitle).AsNoTracking()
                            .Include(t => t.OfficeLocation).AsNoTracking()
                            .Include(t => t.Department).AsNoTracking()
                            .Include("GroupDetails.Group")
                            .ToArray();
                        userPaging.Users = Mapper.Map<List<UserForGrid>>(users);
                    }
                    else
                    {
                        var users = result?
                            .Include(t => t.UserInformation).AsNoTracking()
                            .Include(t => t.JobTitle).AsNoTracking()
                            .Include(t => t.OfficeLocation).AsNoTracking()
                            .Include(t => t.Department).AsNoTracking()
                            .Include("GroupDetails.Group")
                            .ToArray();
                        userPaging.Users = Mapper.Map<List<UserForGrid>>(users);
                    }
                }
                return userPaging;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUsersByFilter", ex.StackTrace);
                throw;
            }
        }

        public IQueryable<Entities.User> FilterUser(UsersFilter filter, IQueryable<Entities.User> result)
        {
            if (!string.IsNullOrEmpty(filter.SearchString))
            {
                result = from user in result
                         where (user.DisplayName.Trim().ToLower().Contains(filter.SearchString.Trim().ToLower()) || user.UserInformation.Email.Trim().ToLower().Contains(filter.SearchString.Trim().ToLower()))
                         select user;
            }
            if (filter.StartDate != null)
            {
                result = from user in result
                         where user.StartDate >= filter.StartDate
                         select user;
            }
            if (filter.Office != Guid.Empty)
            {
                result = from user in result
                         where user.OfficeLocation.Id == filter.Office
                         select user;
            }
            if (filter.Department != Guid.Empty)
            {
                result = from user in result
                         where user.Department.Id == filter.Department
                         select user;
            }
            if (filter.JobTitle != Guid.Empty)
            {
                result = from user in result
                         where user.JobTitle.Id == filter.JobTitle
                         select user;
            }
            if (filter.Gender.HasValue)
            {
                result = from user in result
                         where user.UserInformation.Gender == filter.Gender.Value
                         select user;
            }
            return result;
        }
        public void UpdateAzureUserIdForMissUsers()
        {
            var users = _unitOfWork.UserRepository.Find(x => x.AzureADUserId == Guid.Empty).Include(x => x.UserInformation);
            foreach (var user in users)
            {
                user.AzureADUserId = GetAzureADUserId(user.UserInformation.Email);
                _unitOfWork.UserRepository.Update(user);
            }
            _unitOfWork.Save();
        }
        //TODO: Should devide into Update Leave Track/Update User and Update User Information
        public void UpdateUser(User user)
        {
            try
            {
                var foundUserInformation = _unitOfWork.UserInformationRepository.FindById(user.UserInformation.Id);
                if (foundUserInformation != null)
                {
                    // update UserInformation
                    foundUserInformation.Email = user.UserInformation.Email;
                    foundUserInformation.Phone = user.UserInformation.Phone;
                    foundUserInformation.Birthday = user.UserInformation.Birthday;
                    foundUserInformation.FirstName = user.UserInformation.FirstName;
                    foundUserInformation.LastName = user.UserInformation.LastName;
                    foundUserInformation.Gender = user.UserInformation.Gender;
                    _unitOfWork.Save();
                }

                var foundUser = _unitOfWork.UserRepository.FindById(user.Id);
                if (foundUser != null)
                {
                    var foundStartDate = foundUser.StartDate;
                    // update User
                    foundUser.DisplayName = user.DisplayName;
                    foundUser.Avatar = user.Avatar;
                    foundUser.StartDate = user.StartDate;
                    foundUser.OfficeLocationId = user.OfficeLocationId;
                    foundUser.DepartmentId = user.DepartmentId;
                    foundUser.JobTitleId = user.JobTitleId;
                    foundUser.IsSpUser = user.IsSpUser;
                    foundUser.Modified = DateTime.UtcNow;
                    foundUser.ModifiedBy = user.ModifiedBy;
                    foundUser.ExternalId = user.ExternalId;
                    foundUser.UserStatus = user.UserStatus;
                    _unitOfWork.Save();

                    // check to calculate user leave track
                    if (user.StartDate != foundStartDate)
                    {
                        var leaveTypeQuery = from userRole in _unitOfWork.UserRoleRepository.DbSet()
                                             join leaveType in _unitOfWork.LeaveTypeRepository.DbSet()
                                                 on userRole.AuthoringSiteId equals leaveType.AuthoringSiteId
                                             where userRole.UserId == user.Id && userRole.IsDefault
                                             select leaveType;
                        var leaveTypeList = Mapper.Map<List<LeaveType>>(
                            leaveTypeQuery.Include("LeaveTypeParams.LeaveTypeParamDetails").ToList());

                        foreach (var leaveType in leaveTypeList)
                        {
                            var foundUserLeaveTrack = _unitOfWork.UserLeaveTrackRepository
                                .Find(t => t.UserId == user.Id && t.LeaveTypeId == leaveType.Id)
                                .FirstOrDefault();
                            var newValue = CalculateNumberOfLeaveDaysBySeniorityPolicy(leaveType,
                                user.StartDate.Value, user.UserInformation.Gender);
                            var oldTrack = newValue - foundUserLeaveTrack.TotalLeaveDay;
                            foundUserLeaveTrack.TotalLeaveDay = newValue;
                            foundUserLeaveTrack.RemainLeaveDay = foundUserLeaveTrack.RemainLeaveDay + oldTrack;
                            foundUserLeaveTrack.Modified = DateTime.UtcNow;
                            _unitOfWork.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "UpdateUser", ex.StackTrace);
                throw;
            }
        }

        public List<UserWithUserLeaveTrack> GetUserWithLeaveTrackList(Guid authoringSiteId)
        {
            try
            {
                var userPaging = new UserWithUserLeaveTrackPaging();
                var result = (from user in _unitOfWork.UserRepository.DbSet()
                              join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                              where user.IsActive == true
                                    && userRole.AuthoringSiteId == authoringSiteId
                                    && userRole.IsDefault == true
                              select user)
                    .Include(t => t.UserInformation)
                    .Include(t => t.UserLeaveTracks);
                var userWithUserLeaveTrack = Mapper.Map<List<UserWithUserLeaveTrack>>(result);
                return userWithUserLeaveTrack;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "GetUserLeaveTrack", ex.StackTrace);
                throw;
            }
        }

        public List<UserWithInformationModel> SearchUserInAuthoringSite(SearchUserInAuthoringSite model)
        {
            try
            {
                var userList = (from user in _unitOfWork.UserRepository.DbSet()
                                join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                                where user.IsActive == true
                                      && userRole.AuthoringSiteId == model.AuthoringSiteId
                                      && userRole.IsDefault == true
                                      && user.DisplayName.Contains(model.SearchString)
                                select user)
                    .Include(t => t.UserInformation);
                var result = Mapper.Map<List<UserWithInformationModel>>(userList);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "SearchUserInAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public List<User> SearchNonSPUserInAuthoringSite(SearchUserInAuthoringSite model)
        {
            try
            {
                var userList = (from user in _unitOfWork.UserRepository.DbSet()
                                join userRole in _unitOfWork.UserRoleRepository.DbSet() on user.Id equals userRole.UserId
                                where user.IsSpUser == false
                                      && user.IsActive == true
                                      && userRole.AuthoringSiteId == model.AuthoringSiteId
                                      && userRole.IsDefault == true
                                      && user.DisplayName.Contains(model.SearchString)
                                select user)
                    .Include(t => t.UserInformation);
                var result = Mapper.Map<List<User>>(userList);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "SearchNonSPUserInAuthoringSite", ex.StackTrace);
                throw;
            }
        }

        public bool UpdateDelegateUser(DelegateUsersPost delegateUsers)
        {
            try
            {
                var delegateUserId = _unitOfWork.UserRepository.FindById(delegateUsers.DelegateUserId)?.Id;
                var listDelegateUser = new List<Entities.DelegateUser>();
                if (delegateUserId != null)
                {
                    listDelegateUser = _unitOfWork.DelegateUserRepository.Find(t => t.DelegateId == delegateUserId && t.DelegateForUserId != delegateUserId).ToList();
                }
                foreach (var delegateForUserId in delegateUsers.DelegateForUserId)
                {
                    var existedDelegateUser = listDelegateUser.Find(t => t.DelegateForUserId == delegateForUserId);
                    if (existedDelegateUser == null)
                    {
                        var delegateUser = new Entities.DelegateUser
                        {
                            Id = Guid.NewGuid(),
                            DelegateId = delegateUserId.Value,
                            DelegateForUserId = _unitOfWork.UserRepository.FindById(delegateForUserId).Id,
                            Created = delegateUsers.Created,
                            CreatedBy = delegateUsers.CreatedBy,
                            Modified = delegateUsers.Modified,
                            ModifiedBy = delegateUsers.ModifiedBy,
                        };
                        _unitOfWork.DelegateUserRepository.Add(delegateUser);
                    }
                    else
                    {
                        listDelegateUser.Remove(existedDelegateUser);
                    }
                }
                foreach (var deletedDelegateUser in listDelegateUser)
                {
                    _unitOfWork.DelegateUserRepository.Delete(deletedDelegateUser);
                }
                var existed = _unitOfWork.DelegateUserRepository.Find(t => t.DelegateId == delegateUserId && t.DelegateForUserId == delegateUserId).ToList();
                if (existed == null || existed.Count() == 0)
                {
                    _unitOfWork.DelegateUserRepository.Add(new Entities.DelegateUser
                    {
                        Id = Guid.NewGuid(),
                        DelegateId = delegateUserId.Value,
                        DelegateForUserId = _unitOfWork.UserRepository.FindById(delegateUserId).Id,
                        Created = delegateUsers.Created,
                        CreatedBy = delegateUsers.CreatedBy,
                        Modified = delegateUsers.Modified,
                        ModifiedBy = delegateUsers.ModifiedBy,
                    });
                }
                _unitOfWork.Save();
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "UpdateDelegateUser", ex.StackTrace);
                return false;
            }
        }

        public void RemoveUser(ClientContext ctx, Omnia.HR.Repositories.Entities.User user, string userRole, Guid authoringSiteId, string authoringSiteName)
        {
            try
            {
                if (user.IsSpUser)
                {
                    _spUserRepository.RemoveUserFromSharePointGroup(ctx, user.LoginName, userRole, authoringSiteName);
                }
                RemoveUserFromAuthoringSite(user.Id, authoringSiteId);
                if (!IsUserInAnyAuthoringSite(user.Id))
                {
                    RemoveUserFromDelegateUser(user.Id);
                    RemoveUserFromGroupDetail(user.Id);
                    DeactivateUser(user.Id);
                }
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "RemoveUser", ex.StackTrace);
                throw;
            }
        }

        public UserRole AddUserRole(User newUser, Guid authoringId, int? roleId, bool isAuthoringSite)
        {
            try
            {
                UserRole result;
                if (roleId == null)
                {
                    result = AddDefaultUserRole(newUser.Id, authoringId, isAuthoringSite);
                }
                else
                {
                    var userRole = new UserRolePost
                    {
                        Id = Guid.NewGuid(),
                        UserId = newUser.Id,
                        RoleId = roleId.Value,
                        AuthoringSiteId = authoringId,
                        IsDefault = isAuthoringSite,
                        Created = DateTime.UtcNow,
                        CreatedBy = newUser.Id,
                        Modified = DateTime.UtcNow,
                        ModifiedBy = newUser.ModifiedBy,
                    };


                    result = _userRoleService.AddUserRole(userRole);
                    _unitOfWork.Save();
                }
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddUserRole", ex.StackTrace);
                throw;
            }
        }

        public DelegateUser AddDelegateUser(Guid delegateId, Guid delegateForUserId)
        {
            try
            {
                var existedDelegateUser = _unitOfWork.DelegateUserRepository.Find(t => t.DelegateId == delegateId && t.DelegateForUserId == delegateForUserId).FirstOrDefault();
                if (!existedDelegateUser.IsNotNull())
                {
                    var delegateUser = new Model.Models.DelegateUser
                    {
                        Id = Guid.NewGuid(),
                        DelegateId = delegateId,
                        DelegateForUserId = delegateForUserId,
                        Created = DateTime.UtcNow,
                        Modified = DateTime.UtcNow,
                    };
                    var delegateUserEntity = Mapper.Map<Entities.DelegateUser>(delegateUser);
                    _unitOfWork.DelegateUserRepository.Add(delegateUserEntity);
                    _unitOfWork.Save();

                    return delegateUser;
                }
                return Mapper.Map<DelegateUser>(existedDelegateUser);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddDelegateUser", ex.StackTrace);
                throw;
            }
        }

        public void AddSPUser(DefaultSetting azureCred, UserPost user)
        {
            try
            {
                AzureAuthenticationProvider.AddUser(user, azureCred);
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, DefaultLogTypes.Error, "AddSPUser", ex.StackTrace);
                throw;
            }
        }


        public bool CanCreateSharepointFolder(Guid userId, Guid authoringSiteId)
        {
            var authouringUsers = _userRoleService.GetUserRolesByUserAuthoringSite(userId, authoringSiteId, true);
            return authouringUsers.Any();
        }

        public IList<ExtraEmail> GetEmailFromSPUser(ClientContext ctx)
        {
            return _spUserRepository.GetUserEmails(ctx);
        }

        public bool CreateUserFolder(Guid userId, ClientContext ctx, Guid authoringSiteId)
        {
            var site = ctx.Site;
            var user = _unitOfWork.UserRepository.Find(t => t.Id == userId).FirstOrDefault().UserInformation;
            ctx.Load(site);
            ctx.ExecuteQuery();

            var canCreateFolder = CanCreateSharepointFolder(userId, authoringSiteId);
            var isFolderExist = _spUtilitiesRepository.FolderExists(ctx, HRFolder.DefaultHRDocumentLibraryName, user.Email);
            bool result = false;

            if (canCreateFolder && !isFolderExist)
            {
                OmniaFolder omniaFolder = new OmniaFolder();
                omniaFolder.Name = user.Email;
                omniaFolder.Childs = new System.Collections.Generic.List<OmniaFolder>();
                omniaFolder.Users = new System.Collections.Generic.List<OmniaPrincipal>();

                //Contract folder
                var contractFolder = new OmniaFolder();
                contractFolder.Name = HRFolder.ContractFolderName;
                omniaFolder.Childs.Add(contractFolder);

                //Payslip folder
                var payslipFolder = new OmniaFolder();
                payslipFolder.Name = HRFolder.PayslipFolderName;
                omniaFolder.Childs.Add(payslipFolder);

                ctx.Load(ctx.Web, x => x.ServerRelativeUrl);
                ctx.ExecuteQuery();

                result = _spUtilitiesRepository.CreateFolder(ctx, ctx.Web.ServerRelativeUrl + HRFolder.DefaultHRDocumentLibraryUrl, omniaFolder);
            }
            else
            {
                result = false;
            }
            return result;
        }

        #region Private Methods

        private void RemoveUserFromAuthoringSite(Guid userId, Guid authoringSiteId)
        {
            var userRole = _unitOfWork.UserRoleRepository
                .Find(t => t.UserId == userId && t.AuthoringSiteId == authoringSiteId)
                .FirstOrDefault();

            if (userRole != null)
            {
                _unitOfWork.UserRoleRepository.Delete(userRole.Id);
                _unitOfWork.Save();
            }
        }

        private bool IsUserInAnyAuthoringSite(Guid userId)
        {
            var result = _unitOfWork.UserRoleRepository
                .Find(t => t.UserId == userId)
                .Count();

            return result > 0;
        }

        private void RemoveUserFromDelegateUser(Guid userId)
        {
            var delegateUserList = _unitOfWork.DelegateUserRepository
                .Find(t => t.DelegateId == userId || t.DelegateForUserId == userId)
                .ToList();

            foreach (var delegateUser in delegateUserList)
            {
                _unitOfWork.DelegateUserRepository.Delete(delegateUser.Id);
            }
            _unitOfWork.Save();
        }

        private void RemoveUserFromGroupDetail(Guid userId)
        {
            var groupDetailList = _unitOfWork.GroupDetailRepository
                .Find(t => t.UserId == userId)
                .ToList();

            foreach (var groupDetail in groupDetailList)
            {
                _unitOfWork.GroupDetailRepository.Delete(groupDetail.Id);
            }
            _unitOfWork.Save();
        }

        private void DeactivateUser(Guid userId)
        {
            var foundUser = _unitOfWork.UserRepository.FindById(userId);
            if (foundUser != null)
            {
                foundUser.IsActive = false;
                _unitOfWork.Save();
            }
        }

        private User AddUserDataMapping(ClientContext ctx, UserPost user)
        {
            var newUser = Mapper.Map<User>(user);

            newUser.Id = Guid.NewGuid();
            newUser.LoginName = (user.IsSpUser && user.UserType == UserType.LicensedUser) ? _spUserRepository.EnsureSpLoginName(ctx, user.LoginName) : user.LoginName;
            // newUser.DisplayName = user.IsSpUser ? user.DisplayName : (user.FirstName + " " + user.LastName);
            newUser.DisplayName = user.FirstName + " " + user.LastName;
            newUser.IsActive = true;
            newUser.DepartmentId = user.DepartmentId == Guid.Empty ? null : user.DepartmentId;
            newUser.JobTitleId = user.JobTitleId == Guid.Empty ? null : user.JobTitleId;

            newUser.UserInformation = Mapper.Map<UserInformation>(user);
            newUser.UserInformation.Id = Guid.NewGuid();
            if (user.UserType == UserType.UnliscensedUser || user.UserType == UserType.ExternalUser)
            {
                newUser.UserInformation.Email = user.ExternalEmail;
            }

            return newUser;
        }

        private UserRole AddDefaultUserRole(Guid userId, Guid authoringSiteId, bool isAuthoringSite)
        {
            var role = _unitOfWork.RoleRepository.Find(t => t.Name == RoleType.Member.ToString()).FirstOrDefault();
            var userRolePost = new Model.Models.UserRolePost
            {
                Id = Guid.NewGuid(),
                UserId = userId,
                RoleId = role.Id,
                AuthoringSiteId = authoringSiteId,
                IsDefault = isAuthoringSite,
                Created = DateTime.UtcNow,
                CreatedBy = userId,
                Modified = DateTime.UtcNow,
                ModifiedBy = userId,
            };
            var userRole = _userRoleService.AddUserRole(userRolePost);

            return userRole;
        }

        private Model.Models.UserLeaveYear GetLatestUserLeaveYear(Guid authoringSiteId)
        {
            var result = _unitOfWork.UserLeaveYearRepository.
                Find(t => t.AuthoringSiteId == authoringSiteId)
                .OrderByDescending(t => t.Id).Take(1)
                .FirstOrDefault();

            var item = Mapper.Map<Model.Models.UserLeaveYear>(result);

            return item;
        }

        private List<Model.Models.UserLeaveYear> GetUserLeaveYearByYear(Guid authoringSiteId, int year)
        {
            var result = _unitOfWork.UserLeaveYearRepository.
                Find(t => t.AuthoringSiteId == authoringSiteId && t.Year == year).ToList()
                .Select(x => Mapper.Map<Model.Models.UserLeaveYear>(x)).ToList();



            return result;
        }
        private List<Model.Models.UserLeaveYear> GetUserLeaveYearByRange(Guid authoringSiteId, DateTime? start, DateTime? end)
        {
            var result = new List<Model.Models.UserLeaveYear>();
            if (start.HasValue && !end.HasValue)
            {
                result = _unitOfWork.UserLeaveYearRepository.
                    Find(t => t.AuthoringSiteId == authoringSiteId && t.Year >= start.Value.Year).ToList().Select(x => Mapper.Map<UserLeaveYear>(x)).ToList();
            }
            else if (end.HasValue && !start.HasValue)
            {
                result = _unitOfWork.UserLeaveYearRepository.
                    Find(t => t.AuthoringSiteId == authoringSiteId && t.Year <= end.Value.Year).ToList().Select(x => Mapper.Map<UserLeaveYear>(x)).ToList();
            }
            else
            {
                result = _unitOfWork.UserLeaveYearRepository.
                   Find(t => t.AuthoringSiteId == authoringSiteId && t.Year <= end.Value.Year && t.Year >= start.Value.Year).ToList().Select(x => Mapper.Map<UserLeaveYear>(x)).ToList();
            }

            return result;
        }

        private List<UserWithUserLeaveTrack> MapToUserWithUserLeaveTracks(List<Repositories.Entities.User> userEntities, UserLeaveYear year)
        {
            var users = Mapper.Map<List<User>>(userEntities);
            var results = Mapper.Map<List<UserWithUserLeaveTrack>>(users);
            for (var i = 0; i < users.Count; i++)
            {
                // get singl
                // get user leave track list in latest year
                var userLeaveTrackList = users[i].UserLeaveTracks
                    .Where(t => t.UserLeaveYearId == year.Id)
                    .ToList();
                var mappedUserLeaveTracks = new List<UserLeaveTrack>();
                // get leave type name
                foreach (var userLeaveTrack in userLeaveTrackList)
                {
                    var leaveTypeEntity = _unitOfWork.LeaveTypeRepository
                        .Find(t => t.Id == userLeaveTrack.LeaveTypeId && t.IsActive,
                            includeProperties: "LeaveTypeLocalizations")
                        .FirstOrDefault();
                    if (leaveTypeEntity != null)
                    {
                        var leaveType = Mapper.Map<LeaveType>(leaveTypeEntity);
                        userLeaveTrack.LeaveType = leaveType;
                        userLeaveTrack.LeaveTypeName = leaveType.LeaveTypeLocalizations[0].Name;

                        mappedUserLeaveTracks.Add(userLeaveTrack);
                    }

                }
                // order ascending by name
                results[i].UserLeaveTracks = mappedUserLeaveTracks
                    .OrderBy(t => t.LeaveTypeId)
                    .ToList();
            }

            return results;
        }
        private List<UserWithUserLeaveTrack> MapToUserWithUserLeaveTracks(List<Repositories.Entities.User> userEntities, IList<UserLeaveYear> years, LeavesFilter filter)
        {
            var users = Mapper.Map<List<User>>(userEntities);
            var results = Mapper.Map<List<UserWithUserLeaveTrack>>(users);
            for (var i = 0; i < users.Count; i++)
            {
                //get single leave Request 

                // get user leave track list in latest year
                var userLeaveTrackList = users[i].UserLeaveTracks
                    .Where(t => years.Select(x => x.Id).Contains(t.UserLeaveYearId))
                    .ToList();
                var mappedUserLeaveTracks = new List<UserLeaveTrack>();
                // get leave type name
                foreach (var userLeaveTrack in userLeaveTrackList)
                {
                    var leaveTypeEntity = _unitOfWork.LeaveTypeRepository
                        .Find(t => t.Id == userLeaveTrack.LeaveTypeId && t.IsActive,
                            includeProperties: "LeaveTypeLocalizations")
                        .FirstOrDefault();
                    if (leaveTypeEntity != null)
                    {
                        var leaveType = Mapper.Map<LeaveType>(leaveTypeEntity);
                        userLeaveTrack.LeaveType = leaveType;
                        userLeaveTrack.LeaveTypeName = leaveType.LeaveTypeLocalizations[0].Name;
                        /// Total leave days
                        filter.LeaveType = leaveType;
                        filter.Status = "All";
                        var leaveRequests = _leaveService.GetMyLeaveListByFilter(users[i].Id, 1, int.MaxValue, filter);
                        if (leaveType.Code?.ToLower() == Constants.WorkFromHomeCode)
                        {
                            userLeaveTrack.TotalLeaveDay = 0;
                            var date = new DateTime(filter.StartDate.Value.Year, filter.StartDate.Value.Month, 1);
                            while (date < filter.EndDate.Value)
                            {
                                //6 days per month for people who have 5 years of service or more
                                if (users[i].StartDate.GetValueOrDefault().AddYears(5) < date)
                                {
                                    userLeaveTrack.TotalLeaveDay += 6;
                                }
                                //4 days per month for people who have 2 years of service or more
                                else if (users[i].StartDate.GetValueOrDefault().AddYears(2) < date)
                                {
                                    userLeaveTrack.TotalLeaveDay += 4;
                                }
                                else
                                {
                                    userLeaveTrack.TotalLeaveDay += 0;
                                }
                                date = date.AddMonths(1);
                            }
                        }
                        if (leaveRequests.LeaveRequests != null)
                        {
                            var totalRequest = leaveRequests.LeaveRequests.Where(x => x.Status != ApprovalStep.Rejected && !x.IsRemoved).Sum(x => x.NumberOfDay);
                            /// re-caculate total leave days
                            userLeaveTrack.RemainLeaveDay = userLeaveTrack.TotalLeaveDay - totalRequest;
                        }
                        else
                        {
                            userLeaveTrack.RemainLeaveDay = userLeaveTrack.TotalLeaveDay;
                        }
                        mappedUserLeaveTracks.Add(userLeaveTrack);
                    }

                }
                // order ascending by name
                results[i].UserLeaveTracks = mappedUserLeaveTracks
                    .OrderBy(t => t.LeaveTypeId)
                    .ToList();
            }

            return results;
        }

        private List<UserWithUserLeaveTrack> MapToUserWithUserLeaveTracksForExport(List<Repositories.Entities.User> userEntities, IList<UserLeaveYear> years, LeavesFilter filter)
        {
            var users = Mapper.Map<List<User>>(userEntities);
            var results = Mapper.Map<List<UserWithUserLeaveTrack>>(users);
            for (var i = 0; i < users.Count; i++)
            {
                //get single leave Request 

                // get user leave track list in latest year
                var userLeaveTrackList = users[i].UserLeaveTracks
                    .Where(t => years.Select(x => x.Id).Contains(t.UserLeaveYearId) && (t.LeaveTypeId == 1 || t.LeaveTypeId == 3 || t.LeaveTypeId == 4))
                    .ToList();
                var mappedUserLeaveTracks = new List<UserLeaveTrack>();
                // get leave type name
                foreach (var userLeaveTrack in userLeaveTrackList)
                {
                    var leaveTypeEntity = _unitOfWork.LeaveTypeRepository
                        .Find(t => t.Id == userLeaveTrack.LeaveTypeId && t.IsActive,
                            includeProperties: "LeaveTypeLocalizations")
                        .FirstOrDefault();
                    if (leaveTypeEntity != null)
                    {
                        var leaveType = Mapper.Map<LeaveType>(leaveTypeEntity);
                        userLeaveTrack.LeaveType = leaveType;
                        userLeaveTrack.LeaveTypeName = leaveType.LeaveTypeLocalizations[0].Name;
                        /// Total leave days
                        filter.LeaveType = leaveType;
                        filter.Status = "All";
                        var leaveRequests = _leaveService.GetMyLeaveListByFilter(users[i].Id, 1, int.MaxValue, filter);
                        if (leaveType.Code?.ToLower() == Constants.WorkFromHomeCode)
                        {
                            userLeaveTrack.TotalLeaveDay = 0;
                            var date = new DateTime(filter.StartDate.Value.Year, filter.StartDate.Value.Month, 1);
                            while (date < filter.EndDate.Value)
                            {
                                //6 days per month for people who have 5 years of service or more
                                if (users[i].StartDate.GetValueOrDefault().AddYears(5) < date)
                                {
                                    userLeaveTrack.TotalLeaveDay += 6;
                                }
                                //4 days per month for people who have 2 years of service or more
                                else if (users[i].StartDate.GetValueOrDefault().AddYears(2) < date)
                                {
                                    userLeaveTrack.TotalLeaveDay += 4;
                                }
                                else
                                {
                                    userLeaveTrack.TotalLeaveDay += 0;
                                }
                                date = date.AddMonths(1);
                            }
                        }
                        if (leaveRequests.LeaveRequests != null)
                        {
                            var totalRequest = leaveRequests.LeaveRequests.Where(x => x.Status != ApprovalStep.Rejected && !x.IsRemoved).Sum(x => x.NumberOfDay);
                            /// re-caculate total leave days
                            userLeaveTrack.RemainLeaveDay = userLeaveTrack.TotalLeaveDay - totalRequest;
                        }
                        else
                        {
                            userLeaveTrack.RemainLeaveDay = userLeaveTrack.TotalLeaveDay;
                        }
                        mappedUserLeaveTracks.Add(userLeaveTrack);
                    }

                }
                // order ascending by name
                results[i].UserLeaveTracks = mappedUserLeaveTracks
                    .OrderBy(t => t.LeaveTypeId)
                    .ToList();
            }

            return results;
        }

        private UserWithUserLeaveTrack MapToUserWithUserLeaveTrack(Entities.User userEntity, int yearId)
        {
            var user = Mapper.Map<User>(userEntity);
            var result = Mapper.Map<UserWithUserLeaveTrack>(user);

            // get user leave track list in latest year
            var userLeaveTrackList = user.UserLeaveTracks
                .Where(t => t.UserLeaveYearId == yearId)
                .ToList();
            var mappedUserLeaveTracks = new List<UserLeaveTrack>();
            // get leave type name
            foreach (var userLeaveTrack in userLeaveTrackList)
            {
                var leaveTypeEntity = _unitOfWork.LeaveTypeRepository
                    .Find(t => t.Id == userLeaveTrack.LeaveTypeId && t.IsActive,
                        includeProperties: "LeaveTypeLocalizations")
                    .FirstOrDefault();

                if (leaveTypeEntity != null)
                {
                    var leaveType = Mapper.Map<LeaveType>(leaveTypeEntity);
                    userLeaveTrack.LeaveTypeName = leaveType.LeaveTypeLocalizations[0].Name;
                    userLeaveTrack.LeaveType = leaveType;
                    mappedUserLeaveTracks.Add(userLeaveTrack);
                }
            }
            // order ascending by name
            result.UserLeaveTracks = mappedUserLeaveTracks
                    .OrderBy(t => t.LeaveTypeId)
                    .ToList();
            return result;
        }

        private IQueryable<Entities.User> GetAllUsersContent(Guid tenantId)
        {
            var authouringSiteIds = _unitOfWork.AuthoringSiteRepository.Find(x => x.TenantId == tenantId).Select(x => x.Id);
            var user = _unitOfWork.UserRepository.Find(t => t.IsActive == true && t.UserRoles.Where(x => authouringSiteIds.Contains(x.AuthoringSiteId)).Any())?.Include(t => t.UserInformation);
            return user;
        }

        private Guid GetAzureADUserId(string email)
        {
            try
            {

                var tenantId = ConfigurationManager.AppSettings["AzureTenantId"];
                var graphClient = new GraphServiceClient(new AzureAuthenticationProviderHeader(tenantId));
                using (var task = Task.Run(async () => await graphClient.Users[email].Request().GetAsync()))
                {
                    while (!task.IsCompleted)
                        Thread.Sleep(200);

                    return Guid.Parse(task.Result.Id);
                }
            }
            catch (Exception ex)
            {
                // _logger.LogException(ex, DefaultLogTypes.Error, "GetAzureADUserId", ex.StackTrace);
                return Guid.Empty;
            }
        }
        #endregion
    }
}
